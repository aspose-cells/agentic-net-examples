// Title: How to delete blank rows and columns on the first worksheet with UpdateReference disabled using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete blank rows and columns without updating cell references | preserve formulas when removing empty rows and columns in an .xlsx workbook with Aspose.Cells | disable UpdateReference in DeleteBlankRowsAndColumns for first worksheet Aspose.Cells | remove blank rows and columns from the first sheet of an Excel file using Aspose.Cells .NET | C# example of DeleteBlankRowsAndColumns false UpdateReference Aspose.Cells
// Tags: DeleteBlankRowsAndColumns UpdateReference false Aspose.Cells | preserve formulas while deleting empty rows Aspose.Cells | first worksheet blank row removal C# | blank column deletion without reference update Aspose.Cells | Excel workbook cleanup Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, accesses the first worksheet, deletes all empty rows and columns while keeping existing formulas (by using DeleteBlankRowsAndColumns with UpdateReference set to false or the equivalent DeleteBlankRows/DeleteBlankColumns calls), and saves the cleaned file to a new location.
class Program
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

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            var worksheet = workbook.Worksheets[0];

            // Delete blank rows and columns (preserves formulas)
            worksheet.Cells.DeleteBlankRows();
            worksheet.Cells.DeleteBlankColumns();

            // Ensure the output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
