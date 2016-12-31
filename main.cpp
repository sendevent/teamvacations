#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"

#include <QDebug>
#define LOG qDebug() << Q_FUNC_INFO

QTXLSX_USE_NAMESPACE

// Data rows/columns positions:
static const struct Layout
{
    const int RowMonth = 1;
    const int RowWeek = 2;
    const int RowDay = 3;

    static const int HeadersShiftRows = 2;

    // Column #1 - Employee's names
    // Increase it if you need additional column for some manual data
    static const int DaysShiftColumns = 1;

} Cells;

// Cells styling - fonts, colors, etc:
static const struct Formats
{
    Formats()
    {
        const QColor borderLight = Qt::lightGray;
        const QColor borderDark = Qt::black;

        MonthName.setFontSize(14);
        MonthName.setFontBold(true);
        MonthName.setHorizontalAlignment(Format::AlignLeft);
        MonthName.setVerticalAlignment(Format::AlignTop);
        MonthName.setLeftBorderStyle(Format::BorderThin);
        MonthName.setLeftBorderColor(borderLight);
        MonthName.setBottomBorderStyle(Format::BorderThin);
        MonthName.setBottomBorderColor(borderLight);

        DayName = Format(MonthName);
        DayName.setFontBold(false);
        DayName.setHorizontalAlignment(Format::AlignHCenter);
        DayName.setVerticalAlignment(Format::AlignVCenter);

        WorkDay = Format(DayName);
        WorkDay.setFontSize(10);
        WorkDay.setPatternBackgroundColor(QColor("#fafafa"));

        Weekend = Format(WorkDay);
        Weekend.setFontColor(Qt::white);
        Weekend.setPatternBackgroundColor(QColor("#93CCEA"));
        Weekend.setLeftBorderColor(borderDark);
        Weekend.setBottomBorderColor(borderDark);
        Weekend.setRightBorderStyle(Format::BorderThin);
        Weekend.setRightBorderColor(borderDark);
        Weekend.setTopBorderStyle(Format::BorderThin);
        Weekend.setTopBorderColor(borderDark);

        MonthName.setFontColor(Qt::white);
    }

    Format MonthName;
    Format DayName;
    Format WorkDay;
    Format Weekend;

} CellStyles;

// Simple holder for Year & Month
struct Calendar
{
    Calendar( int y, int m )
        : Year(y)
        , Month(m)
    {}

    const int Year;
    const int Month;

};

static const QStringList MonthsColors = {
    "#EA80FC",
    "#E040FB",
    "#CE93D8",
    "#D500F9",
    "#AA00FF",
    "#BA68C8",
    "#AB47BC",
    "#9C27B0",
    "#8E24AA",
    "#7B1FA2",
    "#6A1B9A",
    "#4A148C",
};

static const QStringList Team = {
    "Alis",
    "Bob",
    "Eva"
};

void generateMonth( Document* doc, const Calendar& calendar )
{
    QDate d( calendar.Year, calendar.Month, 1);

    const int monthDuration = d.daysInMonth();
    // Month's title header - single cell:
    const int lastCell = d.dayOfYear() + Layout::DaysShiftColumns + monthDuration-1;
    const CellRange r( 1, d.dayOfYear() + Layout::DaysShiftColumns, 1, lastCell );
    doc->mergeCells(r, CellStyles.MonthName);

    // Populate days
    for( int i = 1; i <= monthDuration; ++i)
    {
        const int currColumn = d.dayOfYear() + Layout::DaysShiftColumns;
        const Format currDayStyle = d.dayOfWeek() > 5 ? CellStyles.Weekend : CellStyles.WorkDay;

        Format dayNameStyle(currDayStyle);
        dayNameStyle.setFontBold(true);
        doc->write(Cells.RowWeek, currColumn, d.toString("ddd"), dayNameStyle);

        if(1 == i)
        {
            // Add month's title
            Format monthStyle(CellStyles.MonthName);
            monthStyle.setPatternBackgroundColor(MonthsColors.at(d.month()-1));
            doc->write(Cells.RowMonth, currColumn, QLocale().monthName(calendar.Month), monthStyle);
        }

        for(int e = 0; e < Team.size(); ++e)
            doc->write(Cells.RowDay+e, currColumn, d.toString("dd"), currDayStyle);

        d = d.addDays(1);
    }
}

int main(int argc, char **argv)
{
    QCoreApplication app(argc, argv);

    const int targetYear(2017);
    Document xlsx;

    xlsx.addSheet(QString::number(targetYear));
    xlsx.currentWorksheet()->setGridLinesVisible(true);

    // Populate employee's names
    // and detect the longest one
    static const int NamesColumnId(1);
    int maxName(0);
    for(int row = 1; row <= Team.size(); ++row)
    {
        const QString name = Team.at(row-1);
        maxName = qMax(maxName,name.length());
        xlsx.write(row+Layout::HeadersShiftRows, NamesColumnId, name, CellStyles.WorkDay);
    }

    // Adjust names column size
    const CellRange namesCells( 1+Layout::HeadersShiftRows, NamesColumnId, Team.size(), NamesColumnId );
    xlsx.setColumnWidth( namesCells, maxName );

    // Generate days
    for( int month = 1; month <= 12; ++month )
        generateMonth( &xlsx, Calendar(targetYear,month));

    // Adjust days cells size
    const CellRange daysCells( Layout::HeadersShiftRows, Layout::DaysShiftColumns,
                               12, Layout::DaysShiftColumns + QDate(targetYear,1,1).daysInYear() );
    xlsx.setColumnWidth(daysCells, 4 );

    xlsx.saveAs("vacations.xlsx");

    return 0;
}
