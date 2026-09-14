// Title: Export a worksheet‑scoped named range to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, selects a worksheet‑scoped named range, sets it as the print area, and saves only that area as a PDF with Aspose.Cells. | Write a reusable C# method that accepts an input .xlsx path, a named range name, and an output .pdf path, then exports just that range to PDF while validating file and range existence. | Add comprehensive error handling to a C# Aspose.Cells PDF export routine that uses a named range to define the print area, including clear messages for missing workbook or undefined range.
// Common Searches: Aspose.Cells C# export only a named range to PDF | How to set print area from a worksheet scoped named range before PDF conversion in .NET | C# example for saving a specific Excel range as PDF using Aspose.Cells | Export worksheet scoped defined name to PDF with Aspose.Cells for .NET | Save Excel named range as PDF file using Aspose.Cells library
// Tags: Aspose.Cells set print area from named range | export named range to PDF C# | worksheet scoped defined name PDF conversion | Aspose.Cells PDF export specific range | C# Aspose.Cells named range handling

using Aspose.Cells;
using System;
using System.IO;

// The sample loads 'input.xlsx', retrieves the worksheet‑scoped named range 'MyRange', assigns its address to the worksheet's print area, and saves the workbook as 'output.pdf' using Aspose.Cells, with checks for file existence and named‑range availability.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";
            const string rangeName = "MyRange";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the collection of defined names (both workbook‑ and worksheet‑scoped)
            NameCollection definedNames = workbook.Worksheets.Names;
            Name targetName = null;

            if (definedNames != null)
            {
                // Try to get the named range by name using the indexer
                targetName = definedNames[rangeName];
            }

            if (targetName == null)
            {
                Console.WriteLine($"Error: Named range \"{rangeName}\" not found.");
                return;
            }

            // Set the print area of the worksheet to the address of the named range
            sheet.PageSetup.PrintArea = targetName.RefersTo;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF; only the defined print area will be rendered
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
