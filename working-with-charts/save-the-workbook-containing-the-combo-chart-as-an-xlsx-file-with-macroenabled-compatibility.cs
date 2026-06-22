using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the combo chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);

        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);

        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B4"].PutValue(12);
        sheet.Cells["C4"].PutValue(22);

        // Add a combo chart (Column + Line)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // First series as column
        chart.NSeries.Add("B2:B4", true);
        // Second series as line
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries[1].Type = ChartType.Line;

        // Set category (X) axis data
        chart.NSeries.CategoryData = "A2:A4";

        // Enable macros (optional, required for macro‑enabled files)
        workbook.Settings.EnableMacros = true;

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save("ComboChart.xlsm", SaveFormat.Xlsm);
    }
}