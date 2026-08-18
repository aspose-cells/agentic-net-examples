// Title: Copy an Excel worksheet with all conditional formatting using Aspose.Cells for .NET (C#)
// Description: The sample creates a source workbook if missing, loads it, creates an empty destination workbook, copies the first worksheet with Worksheet.Copy, transfers the entire ConditionalFormattings collection, and saves the result to a new file.
// Keywords: Aspose.Cells worksheet copy C# | preserve conditional formatting Aspose.Cells | copy Excel sheet between workbooks | conditional formatting transfer .NET | Excel automation Aspose.Cells | duplicate worksheet with styles | C# Excel conditional formatting copy | Aspose.Cells copy sheet preserving rules
// Common Searches: copy worksheet keep conditional formatting Aspose.Cells C# | how to transfer conditional formatting between Excel workbooks .NET | Aspose.Cells duplicate sheet with all styles | preserve conditional formatting when copying Excel sheet | C# copy Excel worksheet with conditional rules
// Developer Intent: Copy a worksheet from one workbook to another while retaining every conditional formatting rule.
// Use Cases: Generate regional reports by cloning a master sheet that contains color‑coded thresholds. | Migrate legacy worksheets into a new template without losing conditional formatting logic. | Create reusable templates where the same styled sheet is duplicated across multiple output files.
// AI Prompts: Show C# code to copy an Excel worksheet and preserve all conditional formatting using Aspose.Cells. | Explain the steps required to transfer the ConditionalFormattings collection when duplicating a sheet with Aspose.Cells for .NET. | Provide a concise example that copies a worksheet and retains its conditional formatting rules in a new workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetCopy
{
    // The sample creates a source workbook if missing, loads it, creates an empty destination workbook, copies the first worksheet with Worksheet.Copy, transfers the entire ConditionalFormattings collection, and saves the result to a new file.
    public class CopyWorksheetWithConditionalFormatting
    {
        // Entry point required for console application
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
            // Define file paths
            string sourcePath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Ensure source workbook exists; create a simple one if missing
            if (!File.Exists(sourcePath))
            {
                var tempWorkbook = new Workbook();
                var ws = tempWorkbook.Worksheets[0];
                ws.Name = "SampleSheet";
                ws.Cells["A1"].PutValue("Sample Data");
                tempWorkbook.Save(sourcePath);
                Console.WriteLine($"Created placeholder source workbook at '{sourcePath}'.");
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet from each workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Copy the worksheet contents and formats (includes basic formatting)
            destinationSheet.Copy(sourceSheet);

            // Ensure that all conditional formatting rules are also copied
            destinationSheet.ConditionalFormattings.Copy(sourceSheet.ConditionalFormattings);

            // Save the result
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Worksheet copied successfully. Output saved to '{outputPath}'.");
        }
    }
}
