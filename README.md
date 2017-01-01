A simple tool to generate spreadsheet with Person+Calendar data per row - a nice way to track teammates' vacations.

![teamvacarions](https://cloud.githubusercontent.com/assets/2843765/21580358/3b8dc748-d000-11e6-9ebc-219b1b503c8f.png)

It's based on Qt 5 & [QtXlsxWriter](https://github.com/dbzhang800/QtXlsxWriter), the latter one is included as a submodule - don't forget to update it after cloning:

```bash
git submodule init
git submodule update
```

***Notice:***
QtXlsxWriter has an [issue](https://github.com/dbzhang800/QtXlsxWriter/issues/108) with Qt 5.6. For our case is enough just to comment out the problem method's body.
