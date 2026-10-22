// Title: Load an Excel 97‑2003 (.xls) workbook in C# with Aspose.Cells and pre‑load file existence check
// Common Searches: C# Aspose.Cells load .xls file with file existence verification | How to open Excel 97-2003 workbook using Aspose.Cells in .NET | Aspose.Cells Workbook constructor handling missing .xls file | LoadOptions.Version Excel97to2003 example in C# | Reading legacy .xls files with Aspose.Cells and error handling
// Tags: load Excel 97-2003 workbook using Aspose.Cells | Workbook initialization with file existence validation | LoadOptions.Version Excel97to2003 in C# | handling FileNotFoundException in Aspose.Cells | reading legacy .xls format with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that "input.xls" exists, then creates an Aspose.Cells Workbook to load the Excel 97‑2003 file, prints a success message, and catches any exceptions such as a missing file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";

            // Verify that the input file exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook; Aspose.Cells automatically detects the format
            Workbook workbook = new Workbook(inputPath);

            Console.WriteLine("Workbook loaded successfully.");
            // Additional processing can be placed here
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
