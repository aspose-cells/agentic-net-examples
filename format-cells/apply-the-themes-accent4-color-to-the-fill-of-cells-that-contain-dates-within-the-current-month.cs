// Title: Aspose.Cells .NET – Apply Theme Accent4 Background to Dates in the Current Month with Conditional Formatting
// Description: C# example that creates a workbook, inserts sample dates, adds a TimePeriod.ThisMonth conditional formatting rule to range A2:A5, and sets the cell background to the workbook’s Accent4 theme color before saving the file.
// Keywords: Aspose.Cells | C# | conditional formatting | TimePeriod.ThisMonth | theme accent color | Accent4 background | highlight current month dates | Excel automation | cell style | date based formatting
// Common Searches: Aspose.Cells highlight dates this month | apply theme accent color to cells C# | conditional formatting current month Aspose.Cells | set background theme color in Excel using .NET | how to use TimePeriod.ThisMonth with Aspose.Cells
// Developer Intent: Automatically shade cells that contain dates falling in the current month using the workbook’s Accent4 theme color.
// Use Cases: Financial statements where today’s month entries need quick visual identification. | Project timelines that emphasize the current month while preserving corporate theme colors. | Dashboard reports that automatically color‑code current‑month dates for better readability.
// AI Prompts: Show how to also change the font color to white when the Accent4 background is applied. | Provide code that uses Accent2 for dates in the previous month instead of Accent4. | Explain how to extend the Accent4 rule to additional columns such as B and C.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, inserts sample dates, adds a TimePeriod.ThisMonth conditional formatting rule to range A2:A5, and sets the cell background to the workbook’s Accent4 theme color before saving the file.
class ApplyAccent4ToCurrentMonthDates
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data: put some dates in column A
        worksheet.Cells["A1"].PutValue("Date");
        worksheet.Cells["A2"].PutValue(DateTime.Now);                     // today (this month)
        worksheet.Cells["A3"].PutValue(DateTime.Now.AddDays(-10));        // this month
        worksheet.Cells["A4"].PutValue(DateTime.Now.AddMonths(-1));       // last month
        worksheet.Cells["A5"].PutValue(DateTime.Now.AddMonths(1));        // next month

        // Add a conditional formatting collection to the worksheet
        int cfIndex = worksheet.ConditionalFormattings.Add();
        FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

        // Define the range to which the conditional formatting will be applied (A2:A5)
        CellArea area = new CellArea { StartRow = 1, EndRow = 4, StartColumn = 0, EndColumn = 0 };
        fcs.AddArea(area);

        // Add a TimePeriod condition (dates occurring in this month)
        int conditionIndex = fcs.AddCondition(FormatConditionType.TimePeriod);
        FormatCondition fc = fcs[conditionIndex];
        fc.TimePeriod = TimePeriodType.ThisMonth;

        // Apply the theme's Accent4 color to the cell background using a theme color
        fc.Style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.0);

        // Save the workbook
        workbook.Save("Accent4_CurrentMonthDates.xlsx");
    }
}
