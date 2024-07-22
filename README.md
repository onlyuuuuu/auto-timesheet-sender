# Auto Timesheet Sender

A very fun project

Build it first, already a fat JAR
```
./gradlew clean build
```

Then run it
```
java -jar build/libs/auto-timesheet-sender.jar '<base directory>' '<sender email> '<password>' '<recipients>' '<cc>' '<project>'
```

Currently only `@outlook.com` mail domain is supported, no app password, just the actual password would do.

If you just want a "dry run", just add at the end of the command with `--no-email-mode`
