// Title: C# – Remove All External Links from an Excel Workbook with Aspose.Cells
// Description: Loads or creates a workbook, clears every external reference using Workbook.Worksheets.ExternalLinks.Clear(true) so formulas stay valid, and saves a self‑contained file.
// Keywords: Aspose.Cells | C# | remove external links | clear external references | self‑contained Excel | Workbook.Worksheets.ExternalLinks.Clear | Excel automation
// Common Searches: how to delete external links in Excel using Aspose.Cells C# | make Excel file independent of other workbooks | Aspose.Cells clear external references programmatically | remove external hyperlinks from worksheets .NET
// Developer Intent: Eliminate all external references so the saved workbook has no links to other files.
// Use Cases: Prepare a spreadsheet for distribution without exposing source paths. | Maintain formula integrity after stripping external data connections. | Batch‑process multiple workbooks to ensure they are fully self‑contained.
// AI Prompts: Generate C# code that lists every external link in a workbook before removing them with Aspose.Cells. | Show how to replace external‑link formulas with their calculated values when clearing links. | Explain the effect of the Clear(true) argument on formula recalculation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads or creates a workbook, clears every external reference using Workbook.Worksheets.ExternalLinks.Clear(true) so formulas stay valid, and saves a self‑contained file.
    public class RemoveExternalLinksDemo
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists; create a simple workbook if it doesn't.
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    wb.Worksheets[0].Name = "Sheet1";
                    wb.Save(inputPath);
                }

                // Load the workbook from the existing file.
                Workbook workbook = new Workbook(inputPath);

                // Remove all external links, updating formulas where possible.
                workbook.Worksheets.ExternalLinks.Clear(true);

                // Save the workbook; it is now self‑contained with no external links.
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point.
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveExternalLinksDemo.Run();
        }
    }
}
