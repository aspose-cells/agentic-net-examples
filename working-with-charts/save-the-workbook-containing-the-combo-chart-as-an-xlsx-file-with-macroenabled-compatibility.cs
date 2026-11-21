// Title: Create a combo chart and save the workbook as a macro‑enabled XLSX (.xlsm) using Aspose.Cells for .NET
// Common Searches: how to generate a combo chart with column and line series and export as .xlsm using Aspose.Cells C# | Aspose.Cells save workbook with charts to macro enabled Excel file in .NET | C# code to create combo chart and save as macro‑enabled XLSX with Aspose.Cells | export Excel workbook with combo chart to .xlsm format using Aspose.Cells for .NET | Aspose.Cells SaveFormat.Xlsm example for chart workbook
// Tags: Aspose.Cells combo chart export to Xlsm | save workbook as macro-enabled XLSX using SaveFormat.Xlsm | C# create column and line series chart Aspose.Cells | Aspose.Cells chart workbook macro-enabled file | Aspose.Cells .NET chart generation and Xlsm saving

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, builds a combo chart that mixes a column series (Sales) and a line series (Profit), and then saves the workbook as a macro‑enabled XLSX file (ComboChart.xlsm) using Aspose.Cells' SaveFormat.Xlsm.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the combo chart
            // Column A: Categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column B: Sales (Column series)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);

            // Column C: Profit (Line series)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(55);

            // Add a combo chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales and Profit";

            // Add column series (Sales)
            int colSeriesIdx = chart.NSeries.Add("B2:B5", true);
            chart.NSeries[colSeriesIdx].Name = "Sales";

            // Add line series (Profit) and set its chart type to Line
            int lineSeriesIdx = chart.NSeries.Add("C2:C5", true);
            chart.NSeries[lineSeriesIdx].Name = "Profit";
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // NOTE: The IsOnSecondaryAxis property is not available in the current Aspose.Cells version.
            // If secondary axis support is required, upgrade to a newer version where this property exists.

            // Save the workbook as a macro‑enabled XLSX file
            workbook.Save("ComboChart.xlsm", SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
