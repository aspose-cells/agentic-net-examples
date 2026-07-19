// Title: Aspose.Cells .NET – Set Z‑Axis (Series Axis) Number Format to Two Decimal Places in a 3‑D Column Chart
// Description: Creates a workbook, adds sample data, inserts a 3‑D column chart, and formats the Z‑axis (Series axis) tick labels with the numeric pattern "0.00" using the SeriesAxis.TickLabels.NumberFormat property before saving the file.
// Keywords: Aspose.Cells Z axis format | 3D chart series axis number format .NET | custom numeric format chart axis | SeriesAxis.TickLabels.NumberFormat | two decimal places Excel chart
// Common Searches: Aspose.Cells set Z axis number format | format series axis labels in 3D chart .NET | apply 0.00 format to chart axis Aspose | how to change Z‑axis tick label format in Aspose.Cells
// Developer Intent: Apply a two‑decimal‑place numeric format to the Z‑axis (Series axis) tick labels of a 3‑D chart.
// Use Cases: Financial dashboards where the depth axis must display currency values rounded to cents. | Scientific reports that require consistent two‑decimal precision on the Z‑axis of 3‑D visualizations. | Automated Excel generation pipelines that enforce uniform axis formatting across multiple charts.
// AI Prompts: Generate C# code with Aspose.Cells that sets the Z‑axis number format to "0.00" for a 3‑D column chart and saves the workbook. | Explain the relationship between the SeriesAxis property and the Z‑axis in 3‑D charts, and show how to customize its tick label format. | Provide a step‑by‑step tutorial for formatting the series axis of a 3‑D chart to display two decimal places, then verify the result in the exported Excel file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a 3‑D column chart, and formats the Z‑axis (Series axis) tick labels with the numeric pattern "0.00" using the SeriesAxis.TickLabels.NumberFormat property before saving the file.
    class SetZAxisNumberFormat
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for a 3‑D chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(1.2345);
                worksheet.Cells["B3"].PutValue(2.3456);
                worksheet.Cells["B4"].PutValue(3.4567);

                worksheet.Cells["C1"].PutValue("Series2");
                worksheet.Cells["C2"].PutValue(4.5678);
                worksheet.Cells["C3"].PutValue(5.6789);
                worksheet.Cells["C4"].PutValue(6.7890);

                // Add a 3‑D column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply custom numeric format (two decimal places) to Z‑axis tick labels
                // For 3‑D charts the Z‑axis corresponds to the Series axis
                chart.SeriesAxis.TickLabels.NumberFormat = "0.00";

                // Define output file path
                string outputPath = "ZAxisNumberFormatDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
