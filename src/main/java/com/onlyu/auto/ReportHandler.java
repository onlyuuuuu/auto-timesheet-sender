package com.onlyu.auto;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URI;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.Month;
import java.time.Year;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Random;

public class ReportHandler implements Closeable
{
    public static final String _FINAL_OUTPUT_FILE_EXTENSION = "xlsx";
    public static final String _TEMPLATE_FILE_NAME = "template." + _FINAL_OUTPUT_FILE_EXTENSION;
    public static final String _REPORT_FILE_NAME_FORMAT = "TotalTimeSheet-%02d-%s-%d." + _FINAL_OUTPUT_FILE_EXTENSION;
    public static final String _TEMP_REPORT_FILE_NAME_FORMAT = "TotalTimeSheet-%02d-%s-%d-temp." + _FINAL_OUTPUT_FILE_EXTENSION;
    public static final String _DATE_PATTERN = "dd/MM/yyyy";
    public static final DateTimeFormatter _DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern(_DATE_PATTERN);

    public static final Random _RANDOM_INSTANCE = new Random();
    public static final String[][] _RANDOM_CONTENTS = new String[][]
    {
        { "Framework/Tool development, building first URI/API mapping feature, Reverse Proxy Interfaces and implementations.", "Core technology for all upcoming applications" },
        { "More core features development for incoming framework/tool. Designing, implementing and deploying initial landing page and other sections.", "User interfaces" },
        { "Tech stacks researching for feature developments and implementations.", "Research & ad hoc" },
        { "Designing and setting up infrastructure. Considering minimal amount of budget spent on resources", "Research & ad hoc" },
    };

    private final File _file;
    private final File _tempFile;
    private final Path _outputFilePath;
    private final Path _outputTempFilePath;
    private final OPCPackage _package;
    private final Workbook _workbook;
    private final Month _month;
    private final Year _year;
    private final LocalDate _startOfMonth;
    private final LocalDate _endOfMonth;
    private final List<WeekEntry> weekEntries;

    ReportHandler(File file) throws IOException, InvalidFormatException
    {
        LocalDate now = LocalDate.now();
        _file = file;
        _outputFilePath = _file.toPath();
        if (!file.exists())
        {
            ClassLoader classloader = Thread.currentThread().getContextClassLoader();
            InputStream inputStream = classloader.getResourceAsStream(_TEMPLATE_FILE_NAME);
            _package = OPCPackage.open(inputStream);
            _workbook = new XSSFWorkbook(_package);
        }
        else
        {
            _package = OPCPackage.open(_file);
            _workbook = new XSSFWorkbook(_package);
        }
        String reportNameNoExt = _file.getName().substring(0, _file.getName().indexOf("."));
        String[] reportNameSplits = reportNameNoExt.split("-");
        _month = Month.of(Integer.parseInt(reportNameSplits[1]));
        _year = Year.of(Integer.parseInt(reportNameSplits[3]));
        _tempFile = new File(getFinalTempFileAbsolutePath().toString());
        _outputTempFilePath = getFinalTempFileAbsolutePath();
        _startOfMonth = LocalDate.of(_year.getValue(), _month, 1);
        _endOfMonth = _startOfMonth.withDayOfMonth(_startOfMonth.lengthOfMonth());
        weekEntries = new ArrayList<>();
        int weekEntryIndex = 0;
        WeekEntry weekEntry = WeekEntry.of(_startOfMonth, _endOfMonth, _workbook.getSheetAt(0), weekEntryIndex);
        while (weekEntry.getNumberOfDays() > 0)
        {
            weekEntries.add(weekEntry);
            weekEntry = WeekEntry.of(weekEntries.getLast().getEndOfWeek().plusDays(1), _endOfMonth, _workbook.getSheetAt(0), ++weekEntryIndex);
        }
    }

    public static ReportHandler of(File file)
    {
        try
        {
            return new ReportHandler(file);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return null;
    }

    @Override
    public void close() throws IOException
    {
        _workbook.close();
        _package.close();
    }

    public ReportHandler updatePeriodTitle()
    {
        Cell cell = _workbook
            .getSheetAt(0)
            .getRow(1)
            .getCell(0);
        String periodTitle = String.format("%s %d", _month.getDisplayName(TextStyle.FULL, Locale.getDefault()), _year.getValue());
        cell.setCellValue(periodTitle);
        return this;
    }

    public ReportHandler updateStartOfMonth()
    {
        Cell cell = _workbook
            .getSheetAt(0)
            .getRow(3)
            .getCell(5);
        cell.setCellValue(_startOfMonth.format(_DATE_TIME_FORMATTER));
        return this;
    }

    public ReportHandler updateEndOfMonth()
    {
        Cell cell = _workbook
            .getSheetAt(0)
            .getRow(4)
            .getCell(5);
        cell.setCellValue(_endOfMonth.format(_DATE_TIME_FORMATTER));
        return this;
    }

    public ReportHandler updateWeekPeriods()
    {
        for (int i = 0; i < weekEntries.size(); i++)
        {
            weekEntries.get(i)
                .period(String.format("Week %s", i + 1))
                .from(weekEntries.get(i).getStartOfWeek().format(_DATE_TIME_FORMATTER))
                .to(weekEntries.get(i).getEndOfWeek().format(_DATE_TIME_FORMATTER));
        }
        return this;
    }

    public ReportHandler updateContent()
    {
        for (int i = weekEntries.size() - 1; i >= 0; i--)
        {
            int randomInt = _RANDOM_INSTANCE.nextInt(4);
            WeekEntry weekEntry = weekEntries.get(i);
            if (weekEntry.hasContent())
                break;
            if (weekEntry.isFuture())
                continue;
            weekEntry
                .beginAt("18:00")
                .endAt("23:00")
                .totalTime("5h/day = 25h/week")
                .taskDescription(_RANDOM_CONTENTS[randomInt][0])
                .note(_RANDOM_CONTENTS[randomInt][1])
                .totalTimeCalculated("25h");
        }
        return this;
    }

    public ReportHandler markAllAsSent()
    {
        for (int i = weekEntries.size() - 1; i >= 0; i--)
        {
            WeekEntry weekEntry = weekEntries.get(i);
            if (weekEntry.hasBeenSent())
                break;
            if (weekEntry.isCurrent() || weekEntry.isFuture())
                continue;
            weekEntry.markAsSent();
        }
        return this;
    }

    public boolean hasUnsentContent()
    {
        for (int i = weekEntries.size() - 1; i >= 0; i--)
        {
            WeekEntry weekEntry = weekEntries.get(i);
            if (weekEntry.isPast() && weekEntry.hasContent() && !weekEntry.hasBeenSent())
                return true;
        }
        return false;
    }

    public String getFinalSheetName()
    {
        return String.format("%02d-%d", _month.getValue(), _year.getValue());
    }

    public String getFinalFileName()
    {
        return String.format
        (
            _REPORT_FILE_NAME_FORMAT,
            _month.getValue(),
            _month.getDisplayName(TextStyle.FULL, Locale.getDefault()),
            _year.getValue()
        );
    }

    public Path getFinalFileAbsolutePath()
    {
        return Path.of(_file.getParent(), getFinalFileName());
    }

    public String getFinalTempFileName()
    {
        return String.format
        (
            _TEMP_REPORT_FILE_NAME_FORMAT,
            _month.getValue(),
            _month.getDisplayName(TextStyle.FULL, Locale.getDefault()),
            _year.getValue()
        );
    }

    public Path getFinalTempFileAbsolutePath()
    {
        return Path.of(_file.getParent(), getFinalTempFileName());
    }

    public File save() throws IOException
    {
        return write();
    }

    public File write() throws IOException
    {
        try (OutputStream outputStream = new FileOutputStream(_outputTempFilePath.toString()))
        {
            _workbook.setSheetName(0, getFinalSheetName());
            _workbook.write(outputStream);
        }
        File temp = new File(_outputTempFilePath.toString());
        File dest = new File(_outputFilePath.toString());
        if (dest.exists())
            dest.delete();
        temp.renameTo(dest);
        return dest;
    }

    public Month getMonth()
    {
        return _month;
    }

    public String getMonthFullname()
    {
        return _month.getDisplayName(TextStyle.FULL, Locale.getDefault());
    }

    public Year getYear()
    {
        return _year;
    }

    public LocalDate getStartOfMonth()
    {
        return _startOfMonth;
    }

    public LocalDate getEndOfMonth()
    {
        return _endOfMonth;
    }

    public List<WeekEntry> getWeekEntries()
    {
        return weekEntries;
    }

    public File getFile()
    {
        return _file;
    }

    public Path getOutputFilePath()
    {
        return _outputFilePath;
    }
}
