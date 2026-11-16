// Title: Save an Aspose.Cells workbook with a stacked progress‑bar chart to XLSX while preserving cell formulas (C#)
// Common Searches: C# Aspose.Cells save workbook with stacked bar chart and preserve formulas | how to keep formulas when exporting Aspose.Cells chart to XLSX in .NET | generate progress bar chart in Aspose.Cells and write to XLSX without losing formulas | ensure output folder exists before saving Aspose.Cells workbook in C#
// Tags: XLSX export preserving formulas Aspose.Cells | stacked bar progress visualization Aspose.Cells | chart series color customization Aspose.Cells | automatic output directory creation C# Aspose.Cells | cell formula retention on workbook save Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds task data with completed percentages and formulas for remaining values, builds a stacked bar chart styled as a progress bar, ensures the target directory exists, and saves the file as XLSX while keeping all cell formulas intact.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and give it a meaningful name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProgressBar";

            // -------------------------------------------------
            // Populate data for the progress bar
            // -------------------------------------------------
            // Header row
            sheet.Cells["A1"].PutValue("Task");
            sheet.Cells["B1"].PutValue("Completed");
            sheet.Cells["C1"].PutValue("Remaining");

            // Sample tasks
            sheet.Cells["A2"].PutValue("Task 1");
            sheet.Cells["A3"].PutValue("Task 2");
            sheet.Cells["A4"].PutValue("Task 3");

            // Completed percentages (static values)
            sheet.Cells["B2"].PutValue(0.3);   // 30%
            sheet.Cells["B3"].PutValue(0.6);   // 60%
            sheet.Cells["B4"].PutValue(0.9);   // 90%

            // Remaining percentages calculated via formulas (preserved on save)
            sheet.Cells["C2"].Formula = "1-B2";
            sheet.Cells["C3"].Formula = "1-B3";
            sheet.Cells["C4"].Formula = "1-B4";

            // -------------------------------------------------
            // Add a stacked bar chart to act as a progress bar
            // -------------------------------------------------
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 10);
            Chart progressChart = sheet.Charts[chartIndex];
            progressChart.Title.Text = "Progress Bar";

            // Series 1: Completed (green)
            progressChart.NSeries.Add("B2:B4", true);
            progressChart.NSeries[0].Name = "Completed";
            progressChart.NSeries[0].Area.ForegroundColor = Color.Green;

            // Series 2: Remaining (light gray)
            progressChart.NSeries.Add("C2:C4", true);
            progressChart.NSeries[1].Name = "Remaining";
            progressChart.NSeries[1].Area.ForegroundColor = Color.LightGray;

            // -------------------------------------------------
            // Save the workbook to XLSX while preserving all formulas
            // -------------------------------------------------
            string outputPath = "ProgressBarChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
