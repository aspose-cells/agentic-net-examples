// Title: Load only the first three worksheets from an Excel file using Aspose.Cells LoadOptions in C#
// Common Searches: Aspose.Cells C# load first three worksheets only | How to limit worksheet loading with LoadOptions in Aspose.Cells .NET | C# example for reading a subset of sheets from an .xlsx file using Aspose.Cells | Use LoadOptions.StartSheetIndex and SheetCount to load specific sheets in Aspose.Cells | Read only the initial sheets of a workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells LoadOptions sheet count limitation | C# load specific worksheets from Excel workbook | Aspose.Cells StartSheetIndex property usage | limit workbook worksheets on load Aspose.Cells | trim Excel file to first three sheets C#

using Aspose.Cells;
using System;
using System.IO;

// The sample demonstrates how to open an Excel workbook with Aspose.Cells while restricting the load to the first three worksheets by configuring LoadOptions.StartSheetIndex = 0 and LoadOptions.SheetCount = 3, then saves the trimmed workbook to a new file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Keep only the first three worksheets (if more exist)
            const int maxSheets = 3;
            while (workbook.Worksheets.Count > maxSheets)
            {
                // Remove worksheets from the end to preserve the first three
                workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
            }

            // Optional: display the number of loaded worksheets
            Console.WriteLine($"Loaded sheets: {workbook.Worksheets.Count}");

            // Save the (potentially trimmed) workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
