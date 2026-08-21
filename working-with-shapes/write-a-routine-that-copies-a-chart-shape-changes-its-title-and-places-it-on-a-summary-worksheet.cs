// Title: Copy and Rename a Chart Shape to a Summary Worksheet with Aspose.Cells for .NET
// Description: Loads a workbook, ensures a chart exists on the first sheet, copies its ChartShape to a "Summary" worksheet, updates the copied shape's title, and saves the result. Includes fallback chart creation and position control.
// Keywords: Aspose.Cells copy chart shape | C# chart duplication | change chart title Aspose | add chart to summary sheet | ChartShape AddCopy | .NET Excel chart automation
// Common Searches: Aspose.Cells copy chart to another worksheet | C# change chart title after copying | How to add a chart to a Summary sheet with Aspose | Duplicate Excel chart programmatically .NET | Copy first chart and rename it using Aspose.Cells
// Developer Intent: Duplicate the first chart in a workbook, modify its title, and place the copy on a dedicated Summary worksheet using Aspose.Cells for .NET.
// Use Cases: Create a consolidated report that gathers key charts onto a single Summary sheet. | Automate chart replication for presentation decks while reflecting a new context in the title. | Build a workbook template that always includes a pre‑positioned chart on a Summary tab.
// AI Prompts: Generate C# code with Aspose.Cells to copy a ChartShape from one worksheet to another and set a new title. | Explain how to check for existing charts on a sheet and create a default chart if none are found before copying. | Show how to position a copied chart shape at specific rows and columns on the target worksheet using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// Loads a workbook, ensures a chart exists on the first sheet, copies its ChartShape to a "Summary" worksheet, updates the copied shape's title, and saves the result. Includes fallback chart creation and position control.
public class ChartCopyUtility
{
    // Copies the first chart from the first worksheet, changes its title,
    // and places the copied chart shape on a worksheet named "Summary".
    public static void CopyChartToSummary(string sourceFilePath, string outputFilePath)
    {
        try
        {
            // Ensure source file exists; if not, create a minimal workbook.
            if (!File.Exists(sourceFilePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(sourceFilePath);
            }

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(sourceFilePath);

            // Ensure there is at least one worksheet with a chart
            Worksheet sourceSheet = workbook.Worksheets[0];
            if (sourceSheet.Charts.Count == 0)
            {
                // Create a simple chart if none exists (create rule)
                int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sourceSheet.Charts[chartIdx];
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["B3"].PutValue(20);
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";
            }

            // Get the first chart's ChartShape (source shape)
            Chart sourceChart = sourceSheet.Charts[0];
            ChartShape sourceChartShape = sourceChart.ChartObject;

            // Get or create the summary worksheet
            Worksheet summarySheet = workbook.Worksheets["Summary"];
            if (summarySheet == null)
            {
                summarySheet = workbook.Worksheets.Add("Summary");
            }

            // Copy the chart shape to the summary sheet at a desired position
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset)
            Shape copiedShape = summarySheet.Shapes.AddCopy(sourceChartShape, 2, 0, 2, 0);

            // Change the title of the copied shape
            copiedShape.Title = "Summary Chart";

            // Save the workbook (save rule)
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during chart copy operation: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point required for compilation
    public static void Main(string[] args)
    {
        try
        {
            // Example file paths; adjust as needed.
            string sourcePath = "SourceWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            ChartCopyUtility.CopyChartToSummary(sourcePath, outputPath);

            Console.WriteLine("Chart copy operation completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
