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
