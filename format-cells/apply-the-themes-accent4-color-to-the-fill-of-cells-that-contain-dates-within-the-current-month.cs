using System;
using Aspose.Cells;

namespace AsposeCellsAccent4CurrentMonth
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // OPTIONAL: add some sample dates for demonstration
            worksheet.Cells["A1"].PutValue(DateTime.Now);                     // today (this month)
            worksheet.Cells["A2"].PutValue(DateTime.Now.AddMonths(-1));      // last month
            worksheet.Cells["A3"].PutValue(DateTime.Now.AddMonths(1));       // next month
            worksheet.Cells["B1"].PutValue(DateTime.Now.AddDays(-10));       // this month
            worksheet.Cells["B2"].PutValue(DateTime.Now.AddDays(20));        // this month

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = worksheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = worksheet.ConditionalFormattings[cfIndex];

            // Define the range to which the conditional formatting will be applied
            // Here we cover rows 0‑99 and columns 0‑25 (A1‑Z100). Adjust as needed.
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 99,
                StartColumn = 0,
                EndColumn = 25
            };
            fcs.AddArea(area);

            // Add a TimePeriod condition (dates) to the collection
            int conditionIndex = fcs.AddCondition(FormatConditionType.TimePeriod);
            FormatCondition condition = fcs[conditionIndex];

            // Set the time period to "ThisMonth" so only dates in the current month are matched
            condition.TimePeriod = TimePeriodType.ThisMonth;

            // Create a style that uses the theme's Accent4 color for the cell background
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;
            // Apply the theme color Accent4 with no tint (0.0)
            accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, 0.0);
            // Assign the style to the conditional format
            condition.Style = accentStyle;

            // Save the workbook
            workbook.Save("Accent4CurrentMonth.xlsx");
        }
    }
}