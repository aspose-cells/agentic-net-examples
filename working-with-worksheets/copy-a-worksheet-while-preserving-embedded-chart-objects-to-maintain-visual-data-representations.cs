// Title: C# Example: Copy a Worksheet with Embedded Charts Using Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel file, create a sample workbook with a column chart when needed, copy a worksheet (including all chart objects) to a new workbook, keep only the first sheet, and save the result. The code preserves chart data sources and formatting.
// Keywords: Aspose.Cells copy worksheet C# | preserve charts when copying Excel sheet | duplicate worksheet with embedded charts .NET | copy sheet to new workbook Aspose.Cells | Excel chart preservation C# | Aspose.Cells example GitHub | C# Excel worksheet cloning
// Common Searches: how to copy a worksheet and keep its charts in C# | Aspose.Cells copy sheet with charts to new file | preserve embedded chart objects when duplicating Excel sheet | copy first worksheet only using Aspose.Cells .NET | sample code for copying worksheet with charts
// Developer Intent: Copy a worksheet from one workbook to another while retaining all embedded chart objects.
// Use Cases: Generate individual reports by cloning a chart‑rich template sheet into separate workbooks. | Create a backup of a specific sheet that contains visualizations without losing chart formatting. | Extract a single worksheet with its charts from a multi‑sheet workbook for distribution to stakeholders.
// AI Prompts: Provide C# code that copies a specific worksheet containing charts from a source workbook to a new workbook using Aspose.Cells, ensuring the charts are preserved. | Show an Aspose.Cells .NET example that copies only the first worksheet with its embedded charts to a new file and removes any extra sheets. | Explain how to duplicate a worksheet with charts while maintaining data sources and formatting in Aspose.Cells for C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to load an Excel file, create a sample workbook with a column chart when needed, copy a worksheet (including all chart objects) to a new workbook, keep only the first sheet, and save the result. The code preserves chart data sources and formatting.
class CopyWorksheetWithCharts
{
    static void Main()
    {
        try
        {
            string sourcePath = "SourceWithChart.xlsx";

            // Ensure the source file exists; create a simple workbook with a chart if it doesn't.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file '{sourcePath}' not found. Creating a sample workbook with a chart.");

                try
                {
                    Workbook sample = new Workbook();
                    Worksheet ws = sample.Worksheets[0];
                    ws.Name = "Data";

                    // Populate sample data.
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("A");
                    ws.Cells["B2"].PutValue(10);
                    ws.Cells["A3"].PutValue("B");
                    ws.Cells["B3"].PutValue(20);

                    // Add a column chart.
                    int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                    Chart chart = ws.Charts[chartIdx];
                    chart.NSeries.Add("B2:B3", true);
                    chart.NSeries.CategoryData = "A2:A3";
                    chart.Title.Text = "Sample Chart";

                    sample.Save(sourcePath);
                    Console.WriteLine($"Sample workbook created at '{sourcePath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Load the source workbook.
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook to receive the copied worksheet.
            Workbook destinationWorkbook = new Workbook();

            // Copy all worksheets (including charts) from source to destination.
            destinationWorkbook.Copy(sourceWorkbook);

            // Keep only the first worksheet; remove the rest.
            for (int i = destinationWorkbook.Worksheets.Count - 1; i > 0; i--)
            {
                destinationWorkbook.Worksheets.RemoveAt(i);
            }

            // Save the result.
            string destPath = "CopiedWorksheet.xlsx";
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Worksheet copied successfully to '{destPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
