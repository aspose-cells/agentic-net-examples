// Title: Highlight Dates of the Current Month with Theme Accent4 Fill Using Aspose.Cells .NET
// Description: This example creates a workbook, writes sample dates, adds a TimePeriod conditional formatting rule for the "ThisMonth" period, and applies a solid fill using the workbook’s Accent4 theme color to cells that contain dates in the current month. The result is saved as an Excel file.
// Keywords: Aspose.Cells | C# | conditional formatting | TimePeriod | ThisMonth | theme accent color | Accent4 | date highlighting | Excel cell fill | theme color fill
// Common Searches: Aspose.Cells highlight current month dates | apply theme accent color conditional formatting Aspose.Cells | C# set Accent4 fill for date cells | TimePeriod ThisMonth example Aspose.Cells
// Developer Intent: Add a conditional formatting rule that fills cells with dates from the current month using the workbook’s Accent4 theme color.
// Use Cases: Monthly dashboards where dates in the active month stand out automatically. | Scheduling templates that need visual cues for the current month without hard‑coded colors. | Reusable Excel reports that apply a theme‑consistent Accent4 fill to date cells across defined ranges.
// AI Prompts: Generate code to use a different theme accent (e.g., Accent2) while keeping the ThisMonth rule. | Show how to extend the Accent4 fill to multiple columns (A:C) for dates in the current month. | Explain how to add a second conditional formatting rule to the same range without overwriting the Accent4 style.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeAccent4Example
{
    // This example creates a workbook, writes sample dates, adds a TimePeriod conditional formatting rule for the "ThisMonth" period, and applies a solid fill using the workbook’s Accent4 theme color to cells that contain dates in the current month. The result is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data: put some dates in column A
            worksheet.Cells["A1"].PutValue(DateTime.Now);                     // Today (this month)
            worksheet.Cells["A2"].PutValue(DateTime.Now.AddDays(-10));        // Earlier this month
            worksheet.Cells["A3"].PutValue(DateTime.Now.AddMonths(-1));       // Last month (should not be highlighted)
            worksheet.Cells["A4"].PutValue(DateTime.Now.AddMonths(1));        // Next month (should not be highlighted)

            // Add a conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied (e.g., A1:A10)
            CellArea area = new CellArea { StartRow = 0, EndRow = 9, StartColumn = 0, EndColumn = 0 };
            fcs.AddArea(area);

            // Add a TimePeriod condition (dates in the current month)
            int conditionIndex = fcs.AddCondition(FormatConditionType.TimePeriod);
            FormatCondition fc = fcs[conditionIndex];

            // Set the time period to ThisMonth
            fc.TimePeriod = TimePeriodType.ThisMonth;

            // Configure the style: solid fill using the theme's Accent4 color
            fc.Style.Pattern = BackgroundType.Solid;
            fc.Style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.0); // No tint

            // Save the workbook
            workbook.Save("ThemeAccent4CurrentMonth.xlsx");
        }
    }
}
