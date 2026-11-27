// Title: Export an Excel workbook to HTML with visible gridlines while ignoring hidden rows using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export worksheet to HTML with gridlines and hidden rows omitted | How to keep Excel gridlines in HTML output when rows are hidden using Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportGridLines example for hidden rows | Generate HTML from Excel in .NET preserving gridlines but skipping hidden rows | Export hidden rows Excel to HTML without affecting gridline rendering Aspose
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines C# | Excel to HTML conversion with hidden rows omitted | Render Excel gridlines in HTML using Aspose.Cells | Hide rows before HTML export Aspose.Cells | Worksheet.IsGridlinesVisible Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGridLinesDemo
{
    // The sample loads an existing workbook, hides specific rows, ensures worksheet gridlines are visible, configures HtmlSaveOptions with ExportGridLines enabled, and saves the file as HTML. Hidden rows are excluded from the output while gridlines are rendered correctly.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook (it may already contain hidden rows)
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure gridlines are visible in the worksheet itself
                sheet.IsGridlinesVisible = true;

                // Hide specific rows (zero‑based indexes)
                sheet.Rows[2].IsHidden = true; // Row 3
                sheet.Rows[4].IsHidden = true; // Row 5

                // Configure HTML export options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true // Render gridlines in HTML
                    // Hidden rows/columns are omitted by default when they are hidden
                };

                // Save the workbook as HTML using the configured options
                workbook.Save(outputFile, htmlOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
                Console.WriteLine("Gridlines are exported and hidden rows do not affect rendering.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
