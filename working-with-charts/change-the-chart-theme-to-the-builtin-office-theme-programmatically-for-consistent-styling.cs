// Title: Programmatically apply the built‑in Office chart theme to an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart theme to Office style | How to use Aspose.Cells to apply the Office theme to a chart | Apply built‑in Office chart theme with Aspose.Cells .NET API | Aspose.Cells chart color palette Office theme example in C# | Set default Office theme for Excel chart using Aspose.Cells
// Tags: Aspose.Cells chart theme application | ChartColorPaletteType Office usage | Excel chart default Office styling .NET | Programmatic chart theme change Aspose.Cells | Office theme chart formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample sales data, adds a column chart, and demonstrates that the built‑in Office chart theme (the default) is applied. It shows how to set the chart title and save the workbook as an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // The built‑in "Office" theme is the default theme for charts.
            // If a specific theme needs to be applied, Aspose.Cells currently
            // provides color palette changes via ChartColorPaletteType.
            // Here we keep the default styling.

            // Optional: give the chart a title
            chart.Title.Text = "Quarterly Sales";

            // Save the workbook
            string outputPath = "ChartWithOfficeTheme.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
