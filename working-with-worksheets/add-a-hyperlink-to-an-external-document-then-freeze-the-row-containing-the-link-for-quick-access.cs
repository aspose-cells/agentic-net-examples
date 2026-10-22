// Title: Create an Excel file with a hyperlink to a local PDF in cell A1 and freeze the top row using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that generates a workbook, places a clickable link to a local PDF in cell A1, and saves the file. | Show how to use Aspose.Cells Hyperlinks.Add together with FreezePanes to add an external document link and keep the first row static. | Provide a C# snippet that verifies a PDF file exists before adding a hyperlink to cell A1 and then freezes the worksheet’s header row.
// Common Searches: Aspose.Cells C# add hyperlink to local PDF and freeze first row | How to keep header row visible after inserting a hyperlink with Aspose.Cells | C# example for checking file existence before creating Excel hyperlink using Aspose.Cells
// Tags: add hyperlink to cell Aspose.Cells C# | freeze top row Aspose.Cells | PDF file hyperlink Aspose.Cells | check PDF presence before hyperlink Aspose.Cells | create workbook with hyperlink Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program creates a new workbook, writes "Open External Document" in cell A1, adds a hyperlink to a local PDF file, freezes the first row to keep the link visible while scrolling, and saves the workbook as HyperlinkAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a descriptive text in cell A1
            Cell linkCell = sheet.Cells["A1"];
            linkCell.PutValue("Open External Document");

            // Path to the external document
            string targetFile = @"C:\Docs\Report.pdf";

            // Verify that the target file exists to avoid FileNotFoundException
            if (!File.Exists(targetFile))
            {
                Console.WriteLine($"Warning: Target file not found: {targetFile}");
            }

            // Add a hyperlink to the local file (covers a single cell)
            // The Add method expects the address as a string.
            sheet.Hyperlinks.Add(0, 0, 1, 1, targetFile);

            // Freeze the first row so the hyperlink stays visible while scrolling
            sheet.FreezePanes(1, 0, 0, 0);

            // Save the workbook
            string outputPath = "HyperlinkAndFreeze.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
