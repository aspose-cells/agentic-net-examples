// Title: Aspose.Cells for .NET – Set Z‑Axis (Series Axis) Tick Labels to Percentages in a 3‑D Column Chart (C#)
// Description: Creates a workbook, adds sample data, inserts a 3‑D column chart, and formats the Z‑axis (SeriesAxis) tick labels with the number format "0%" so the values appear as percentages before saving the file.
// Keywords: Aspose.Cells Z axis percentage | SeriesAxis tick label format C# | 3D column chart axis formatting | C# Aspose.Cells chart percentage labels | Excel 3D chart series axis format
// Common Searches: format Z axis as percent Aspose.Cells | SeriesAxis number format example C# | 3D chart axis percentage Aspose.Cells .NET | how to show percentages on Z axis in Excel chart using Aspose | C# set series axis tick labels to 0% in Aspose.Cells
// Developer Intent: Apply a percentage number format to the Z‑axis (SeriesAxis) tick labels of a 3‑D chart.
// Use Cases: Display relative values on the Z‑axis of a 3‑D column chart as percentages for clearer comparison. | Generate Excel reports with 3‑D charts where the series axis labels follow business‑standard percentage formatting. | Automate workbook creation that requires percentage‑formatted Z‑axis labels for dashboards or presentations.
// AI Prompts: Give C# code that sets the SeriesAxis tick label format to a percentage in an Aspose.Cells 3‑D column chart. | Explain step‑by‑step how to format the Z‑axis of a 3‑D chart as percentages using Aspose.Cells for .NET. | Show an example that saves an Excel file with a 3‑D column chart where the Z‑axis labels display 10%, 25%, and 50%.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a 3‑D column chart, and formats the Z‑axis (SeriesAxis) tick labels with the number format "0%" so the values appear as percentages before saving the file.
class SetZAxisLabelAsPercentage
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a 3‑D chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(0.10);   // 10%
            sheet.Cells["B3"].PutValue(0.25);   // 25%
            sheet.Cells["B4"].PutValue(0.50);   // 50%

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // For 3‑D charts the Z‑axis is represented by SeriesAxis
            chart.SeriesAxis.TickLabels.NumberFormat = "0%";

            // Prepare output folder if it does not exist
            string outputPath = "ZAxisPercentage.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
