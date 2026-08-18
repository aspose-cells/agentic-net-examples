// Title: Copy a Worksheet with Embedded Charts using Aspose.Cells for .NET
// Description: Demonstrates how to duplicate a worksheet while preserving all embedded chart objects with Aspose.Cells. The sample creates a workbook with a column chart if none exists, copies the sheet using Workbook.Worksheets.AddCopy, renames the copy, and saves the result.
// Keywords: Aspose.Cells copy worksheet | duplicate sheet with charts | preserve chart objects .NET | Workbook.Worksheets.AddCopy | C# Aspose.Cells chart copy | copy worksheet including charts | Aspose.Cells sample workbook with chart
// Common Searches: how to copy a worksheet and keep charts Aspose.Cells | Aspose.Cells AddCopy preserve chart objects | duplicate Excel sheet with embedded charts using C# | copy worksheet with charts to new workbook Aspose.Cells | Aspose.Cells copy sheet including charts example
// Developer Intent: Programmatically copy an existing worksheet and retain every embedded chart without additional processing.
// Use Cases: Generate a reporting template: create a chart once, then copy the sheet to produce multiple reports with identical visualizations. | Perform scenario analysis on a client‑provided workbook by duplicating chart‑rich sheets for what‑if calculations. | Automatically build a sample workbook with a chart when the source file is missing, then copy the sheet to create a ready‑to‑use template.
// AI Prompts: Write C# code with Aspose.Cells that copies a worksheet containing charts and ensures the charts appear in the copied sheet. | Explain whether Workbook.Worksheets.AddCopy automatically copies chart objects or if extra steps are required. | Provide a step‑by‑step tutorial to create a workbook with a column chart, then duplicate that worksheet while preserving the chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for ChartType and Chart classes

namespace AsposeCellsWorksheetCopyWithCharts
{
    // Demonstrates how to duplicate a worksheet while preserving all embedded chart objects with Aspose.Cells. The sample creates a workbook with a column chart if none exists, copies the sheet using Workbook.Worksheets.AddCopy, renames the copy, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourcePath = "SourceWithCharts.xlsx";

                Workbook sourceWorkbook;

                // Ensure the source file exists; if not, create a sample workbook with a chart
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found. Creating a sample workbook with a chart.");

                    sourceWorkbook = new Workbook();
                    Worksheet ws = sourceWorkbook.Worksheets[0];
                    ws.Name = "DataSheet";

                    // Populate sample data
                    ws.Cells["A1"].PutValue(1);
                    ws.Cells["A2"].PutValue(2);
                    ws.Cells["A3"].PutValue(3);
                    ws.Cells["B1"].PutValue(2);
                    ws.Cells["B2"].PutValue(4);
                    ws.Cells["B3"].PutValue(6);

                    // Add a simple column chart
                    int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = ws.Charts[chartIdx];
                    chart.NSeries.Add("A1:B3", true);

                    // Save the generated source workbook
                    sourceWorkbook.Save(sourcePath);
                }
                else
                {
                    // Load the existing workbook
                    sourceWorkbook = new Workbook(sourcePath);
                }

                // Name of the worksheet to be copied
                string sheetToCopy = "DataSheet";

                // Copy the specified worksheet within the same workbook (includes charts)
                int copiedIndex = sourceWorkbook.Worksheets.AddCopy(sheetToCopy);

                // Rename the copied worksheet
                Worksheet copiedSheet = sourceWorkbook.Worksheets[copiedIndex];
                copiedSheet.Name = sheetToCopy + "_Copy";

                // Save the workbook containing both original and copied worksheets
                string outputPath = "WorkbookWithCopiedSheet.xlsx";
                sourceWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
