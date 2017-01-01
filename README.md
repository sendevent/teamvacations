A simple tool to generate spreadsheet with Person+Calendar data per row - a nice way to track teammates' vacations.

![teamvacarions](https://cloud.githubusercontent.com/assets/2843765/21580330/a48541b0-cffe-11e6-8838-416deea79ca4.png)

It's based on Qt 5 & [QtXlsxWriter](https://github.com/dbzhang800/QtXlsxWriter), the latter one is included as a submodule - don't forget to update it after cloning:

```bash
git submodule init
git submodule update
```

***Notice:***
QtXlsxWriter has an [issue](https://github.com/dbzhang800/QtXlsxWriter/issues/108) with Qt 5.6. For our case is enough just to comment out the problem method's body.
