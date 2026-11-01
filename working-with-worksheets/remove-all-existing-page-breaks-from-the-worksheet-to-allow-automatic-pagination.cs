// Title: Remove all horizontal and vertical page breaks from an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that clears both HorizontalPageBreaks and VerticalPageBreaks collections of a worksheet and saves the workbook. | Show how to programmatically delete every page break in the first sheet of an existing .xlsx file using Aspose.Cells, creating the output folder if needed.
// Common Searches: C# Aspose.Cells remove all page breaks from a worksheet before saving | how to clear horizontal and vertical page breaks in an Excel file using Aspose.Cells .NET | Aspose.Cells delete page breaks to enable automatic pagination in .xlsx | remove page breaks from first worksheet programmatically with Aspose.Cells | Aspose.Cells pagination without manual page breaks C# example
// Tags: clear worksheet page breaks Aspose.Cells | Aspose.Cells HorizontalPageBreaks collection | Aspose.Cells VerticalPageBreaks collection | automatic pagination Excel Aspose.Cells | delete page breaks before workbook save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, clears its HorizontalPageBreaks and VerticalPageBreaks collections, ensures the output directory exists, and saves the modified workbook to a new file.
class RemovePageBreaks
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired worksheet index/name)
            Worksheet sheet = workbook.Worksheets[0];

            // Remove all horizontal page breaks (use HorizontalPageBreaks for newer API versions)
            sheet.HorizontalPageBreaks.Clear();

            // Remove all vertical page breaks (use VerticalPageBreaks for newer API versions)
            sheet.VerticalPageBreaks.Clear();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the page breaks removed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
