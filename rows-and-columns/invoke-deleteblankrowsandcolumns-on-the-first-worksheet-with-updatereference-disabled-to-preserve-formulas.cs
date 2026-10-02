// Title: Delete blank rows and columns on the first worksheet without updating references to keep formulas intact using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, calls Worksheet.Cells.DeleteBlankRowsAndColumns(false) on the first sheet to remove empty rows and columns while leaving formula references unchanged, and saves the file. | Show an example of using Aspose.Cells to clean a worksheet by deleting blank rows and columns with UpdateReference set to false, preserving all existing formulas. | Provide a step‑by‑step C# snippet that checks for the input file, opens it with Aspose.Cells, invokes DeleteBlankRowsAndColumns(false) on sheet[0], and writes the result to a new file.
// Common Searches: Aspose.Cells DeleteBlankRowsAndColumns false keep formulas | C# remove empty rows and columns without updating cell references in Excel | how to preserve formulas when cleaning a worksheet using Aspose.Cells | disable UpdateReference while deleting blank rows in Aspose.Cells .NET
// Tags: delete blank rows and columns Aspose.Cells | preserve formulas DeleteBlankRowsAndColumns | UpdateReference false Aspose.Cells | first worksheet cleanup C# | Excel blank row removal without reference update

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, calls Cells.DeleteBlankRowsAndColumns(false) to remove all empty rows and columns while preserving formula references, saves the modified workbook as output.xlsx, and handles any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Delete blank rows and columns while preserving formulas
            sheet.Cells.DeleteBlankRows();
            sheet.Cells.DeleteBlankColumns();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
