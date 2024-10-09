package com.onlyu.auto;

import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.eclipse.angus.mail.util.MailConnectException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.Month;
import java.time.Year;
import java.time.format.TextStyle;
import java.util.*;

public class AutoUpdater
{
    public static final String _HOST = "smtp.gmail.com";
    public static final String _PORT = "465";
    public static final int _RETRY_ATTEMPTS_THRESHOLD = 120;
    public static final String _SUBJECT = "[%s] Andy - Timesheet - %s - %s";
    public static final String _CONTENT =
    """
        Dear everyone,

        Attached in this email is the up-to-date working timesheet for %s - %s.

        Thank you for your support. If you need anything, please let me know.\s

        Best regards,

        Anh Nguyen (Andy)
    """;

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException, MessagingException, InterruptedException
    {
        System.out.printf("Started the timesheet updater process...\n");

        String baseDirPathArg = args[0];
        Path baseDirPath = Path.of(baseDirPathArg);

        String senderArg = args[1];
        String passwordArg = args[2];
        String recipientsArg = args[3];
        String ccRecipientsArg = args[4];
        String projectNameArg = args[5];
        boolean noEmailMode = (args.length > 6 && "--no-email-mode".equals(args[6])) || (args.length > 7 && "--no-email-mode".equals(args[7]));
        boolean dryRunMode = (args.length > 6 && "--dry-run".equals(args[6])) || (args.length > 7 && "--dry-run".equals(args[7]));
        noEmailMode = dryRunMode ? true : noEmailMode;

        System.out.printf("Base directory path provided: [%s]\n", baseDirPathArg);
        File baseDir = new File(baseDirPathArg);
        if (!baseDir.exists())
            throw new FileNotFoundException("Base directory does not exist: " + baseDir.getAbsolutePath());
        System.out.printf("Verified that base directory [%s] exists\n", baseDirPathArg);

        // Preload some values
        LocalDate now = LocalDate.now();
        LocalDate lastDayOfThePreviousMonth = now.withDayOfMonth(1).minusDays(1);
        Month currentMonth = now.getMonth();
        Month previousMonth = lastDayOfThePreviousMonth.getMonth();
        Year currentYear = Year.of(now.getYear());
        Year previousYear = Year.of(lastDayOfThePreviousMonth.getYear()); // This maybe not the actual previous year

        // Previous month
        File previousMonthReport = new File(Path.of
        (
            baseDirPathArg,
            String.format
            (
                ReportHandler._REPORT_FILE_NAME_FORMAT,
                previousMonth.getValue(),
                previousMonth.getDisplayName(TextStyle.FULL, Locale.getDefault()),
                previousYear.getValue()
            )
        ).toString());
        // Current month
        File currentMonthReport = new File(Path.of
        (
            baseDirPathArg,
            String.format
            (
                ReportHandler._REPORT_FILE_NAME_FORMAT,
                currentMonth.getValue(),
                currentMonth.getDisplayName(TextStyle.FULL, Locale.getDefault()),
                currentYear.getValue()
            )
        ).toString());
        // Verify
        System.out.printf("Previous month report: %s\n", previousMonthReport.getAbsolutePath());
        System.out.printf("Current month report:  %s\n", currentMonthReport.getAbsolutePath());
        List<File> reports = Arrays.asList(previousMonthReport, currentMonthReport);
        List<File> toBeSentReports = new ArrayList<>();

        // Real processing
        StringBuilder finalUnsentContentsTitle = new StringBuilder();
        for (File report : reports)
        {
            // It's not existed yet, then create it
            if (!report.exists())
            {
                System.out.printf("Current month report %s not found. Creating a new report from scratch now...\n", currentMonthReport.getName());
                try (ReportHandler handler = ReportHandler.of(report, baseDirPath))
                {
                    report = handler
                        .updatePeriodTitle()
                        .updateStartOfMonth()
                        .updateEndOfMonth()
                        .updateWeekPeriods()
                        .updateContent()
                        .save();
                    System.out.printf("Created report %s is assigned to month: %d\n", report.getName(), handler.getMonth().getValue());
                    System.out.printf("Created report %s is assigned to year:  %d\n", report.getName(), handler.getYear().getValue());
                    System.out.printf("Created report %s start date:           %s\n", report.getName(), handler.getStartOfMonth().format(ReportHandler._DATE_TIME_FORMATTER));
                    System.out.printf("Created report %s end date:             %s\n", report.getName(), handler.getEndOfMonth().format(ReportHandler._DATE_TIME_FORMATTER));
                }
            }

            // Time to do some processing...
            // 1. Load the report or reports
            // 2. If there are any single past content that has not been sent, then sent it
            // 3. If there are not, then just let it be, next week task will handle that

            // Load the report
            try (ReportHandler handler = ReportHandler.of(report, baseDirPath))
            {
                handler
                    .updateContent()
                    .save();
                StringBuilder sb = new StringBuilder();
                for (WeekEntry entry : handler.getWeekEntries())
                {
                    if (entry.isPast() && entry.hasContent() && !entry.hasBeenSent())
                        sb.append(entry.getPresentableIndex() + ", ");
                }
                if (sb.isEmpty())
                    continue;
                sb = sb
                    .delete(sb.length() - 2, sb.length())
                    .insert(0, String.format("%s Week ", handler.getMonthFullname()));
                System.out.printf("%s in %s have unsent content!\n", sb, report.getName());
                finalUnsentContentsTitle.append(sb).append(" & ");
                toBeSentReports.add(report);
            }
        }

        if (finalUnsentContentsTitle.isEmpty() && !dryRunMode)
        {
            System.out.printf("No need to send anything at the moment.\n");
            return;
        }

        finalUnsentContentsTitle = !finalUnsentContentsTitle.isEmpty()
            ? finalUnsentContentsTitle.delete(finalUnsentContentsTitle.length() - 3, finalUnsentContentsTitle.length())
            : finalUnsentContentsTitle;
        System.out.printf("Sending reports for %s now...\n", finalUnsentContentsTitle);

        if (args[1] == null || args[1].isBlank())
            throw new RuntimeException("No SMTP Outlook password was provided!");

        // Mail props
        Properties properties = new Properties();
        properties.put("mail.smtp.host", _HOST);
        properties.put("mail.smtp.port", _PORT);
        properties.put("mail.smtp.auth", "true");
        properties.put("mail.smtp.socketFactory.port", "465");
        properties.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        // Core objects
        Session session = Session.getDefaultInstance(properties, new Authenticator()
        {
            @Override
            protected PasswordAuthentication getPasswordAuthentication()
            {
                return new PasswordAuthentication(senderArg, passwordArg);
            }
        });
        MimeMessage message = new MimeMessage(session);
        // Prepare message
        message.setFrom(new InternetAddress(senderArg));
        message.addRecipient(Message.RecipientType.TO, new InternetAddress(recipientsArg));
        message.addRecipients(Message.RecipientType.CC, InternetAddress.parse(ccRecipientsArg));
        message.setSubject(String.format(_SUBJECT, projectNameArg, finalUnsentContentsTitle, currentYear.getValue()));
        // Body
        BodyPart messageBodyPart = new MimeBodyPart();
        messageBodyPart.setText(String.format(_CONTENT, finalUnsentContentsTitle, currentYear.getValue()));
        // Attachments
        List<MimeBodyPart> attachmentParts = new ArrayList<>();
        for (File report : toBeSentReports)
        {
            MimeBodyPart attachmentPart = new MimeBodyPart();
            attachmentPart.attachFile(report);
            attachmentParts.add(attachmentPart);
        }
        // Assemble
        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(messageBodyPart);
        for (MimeBodyPart bodyPart : attachmentParts)
            multipart.addBodyPart(bodyPart);
        // Integrate
        message.setContent(multipart);
        // Check for internet connection
        System.out.printf("Checking mail server connectivity...\n");
        try (Transport transport = message.getSession().getTransport())
        {
            int attempt = 0;
            while (attempt <= _RETRY_ATTEMPTS_THRESHOLD)
            {
                try
                {
                    transport.connect(senderArg, passwordArg);
                    if (transport.isConnected())
                    {
                        System.out.printf("Acquired connection to mail server. Confirmed that internet connectivity is active!\n");
                        break;
                    }
                }
                catch (MailConnectException e)
                {
                    System.err.printf("Might be internet connection issue, error: %s!\n", e.getMessage());
                }
                ++attempt;
                System.out.printf("Attempting to connect to mail server %s:%s. Number of efforts so far: %d\n", _HOST, _PORT, attempt);
                Thread.sleep(1000); // Go back and try again
            }
            if (attempt > _RETRY_ATTEMPTS_THRESHOLD)
            {
                System.out.printf("No internet connection, probably, nothing is going out at the moment...\n");
                return;
            }
        }
        catch (Exception e)
        {
            System.err.printf("Exception occurred while checking for mail server connectivity\n");
            e.printStackTrace();
        }
        // Are we allow to send it?
        if (noEmailMode)
        {
            System.out.printf("Email sending feature is disabled! Nothing is going out!\n");
            return;
        }
        // Send it!
        System.out.printf("Sending email now...\n");
        Transport.send(message);
        System.out.printf("Email sent!\n");
        // Mark all week entries as sent!
        System.out.printf("Marking week entries as sent...\n");
        for (File report : toBeSentReports)
        {
            try (ReportHandler handler = ReportHandler.of(report, baseDirPath))
            {
                handler
                    .markAllAsSent()
                    .save();
            }
        }
        System.out.printf("Marked week entries as sent!\n");
    }

}
