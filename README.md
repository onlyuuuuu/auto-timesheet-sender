# Auto Timesheet Sender

A very fun project

Build it first, already a fat JAR
```
./gradlew clean build
```

Then run it
```
java -jar build/libs/auto-timesheet-sender.jar '<base directory>' '<sender email>' '<password (should be an app password>' '<recipients>' '<cc>' '<project>'
```

Currently only `@outlook.com` mail domain is supported, no app password, just the actual password would do.

If you just want a "dry run", just add at the end of the command with `--no-email-mode` or `--dry-run`

By the way, difference between a `--no-email-mode` and `--dry-run` is that `--no-email-mode` only disable email sending if there are contents to be sent. With `--dry-run` email sending will be forced even with empty contents. Ultimately, at the end of the day, `--dry-run` is basically a `--no-email-mode` with force sending empty contents. Hope that makes sense!

From 9th October 2024, this program has been fixed to only works on gmail and gmail app password.
