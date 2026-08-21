// Title: Enable Page Break Preview for All Worksheets with Aspose.Cells for .NET (C#)
// Description: This example shows how to create or load an Aspose.Cells workbook, turn on the Page Break Preview view for every worksheet by setting the IsPageBreakPreview property, and save the file so it opens ready for print layout review.
// Keywords: Aspose.Cells page break preview C# | set IsPageBreakPreview all worksheets | Excel print layout Aspose.Cells | .NET workbook view settings | enable page break view programmatically
// Common Searches: Aspose.Cells enable page break preview for each sheet | C# set IsPageBreakPreview property on all worksheets | How to turn on page break view in generated Excel file using Aspose.Cells | Prepare Excel workbook for printing with Aspose.Cells .NET
// Developer Intent: Activate Page Break Preview on every worksheet before saving the workbook.
// Use Cases: Generate reports that display pagination boundaries automatically when opened in Excel. | Create templates where users can instantly see page breaks for consistent printing. | Automate batch exports that require a printable layout without manual view adjustments.
// AI Prompts: Write C# code using Aspose.Cells to enable page break preview for all worksheets and then export the workbook to PDF. | Show how to toggle IsPageBreakPreview based on a configuration flag passed to a method. | Explain how to programmatically confirm that Page Break Preview is active when the workbook is opened in Excel.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example shows how to create or load an Aspose.Cells workbook, turn on the Page Break Preview view for every worksheet by setting the IsPageBreakPreview property, and save the file so it opens ready for print layout review.
    public class ApplyPageBreakPreview
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a workbook with a default worksheet

            // If you need to load an existing workbook, uncomment the following lines
            // string inputPath = "input.xlsx";
            // if (File.Exists(inputPath))
            // {
            //     workbook = new Workbook(inputPath);
            // }
            // else
            // {
            //     Console.WriteLine($"Input file not found: {inputPath}");
            //     return;
            // }

            // Enable Page Break Preview for every worksheet in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.IsPageBreakPreview = true;
            }

            // Save the workbook with the updated view settings
            string outputPath = "Output_PageBreakPreview.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
