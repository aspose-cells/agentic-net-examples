// Title: How to clear only the contents of a specific named range in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, retrieves the named range "ReportData", clears its cell values without deleting the range, and saves the updated file. | Show how to use Aspose.Cells Range.ClearContents in C# to empty a named range while preserving its definition and formatting.
// Common Searches: Aspose.Cells C# clear values of a named range but keep the range definition | C# remove data from Excel named range using Aspose.Cells without deleting the name | How to use Range.ClearContents for a named range in Aspose.Cells .NET | Preserve named range while clearing its cells with Aspose.Cells in C# | Clear contents of ReportData named range in Excel using Aspose.Cells library
// Tags: Aspose.Cells Range.ClearContents C# | remove cell values from named range Aspose.Cells | retain named range definition while clearing cells Aspose.Cells | Excel .xlsx workbook modification Aspose.Cells .NET | named range handling Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads 'input.xlsx' with Aspose.Cells, checks for a named range called 'ReportData', obtains its Range object, calls ClearContents to delete all cell values while keeping the named range intact, and saves the result to 'output.xlsx'. It includes file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ReportData"
            Name namedRange = workbook.Worksheets.Names["ReportData"];
            if (namedRange != null)
            {
                // Get the actual cell range and clear only its contents
                Aspose.Cells.Range range = namedRange.GetRange();
                range.ClearContents();
            }
            else
            {
                Console.WriteLine("Named range 'ReportData' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
