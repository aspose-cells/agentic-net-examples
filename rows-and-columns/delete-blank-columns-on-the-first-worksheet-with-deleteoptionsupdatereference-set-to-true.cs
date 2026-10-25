// Title: How to delete blank columns in the first worksheet of an Excel file using Aspose.Cells for .NET with UpdateReference enabled
// Common Searches: Aspose.Cells C# delete empty columns and keep formulas updated | Delete blank columns on first sheet while updating cell references Aspose.Cells | C# example for Worksheet.Cells.DeleteBlankColumns with DeleteOptions.UpdateReference | Remove all blank columns from a workbook using Aspose.Cells and preserve formulas | How to use DeleteOptions.UpdateReference when cleaning up blank columns in .NET
// Tags: Aspose.Cells DeleteBlankColumns UpdateReference | C# delete blank columns first worksheet | preserve formula references Aspose.Cells | Excel workbook cleanup using Aspose.Cells | DeleteOptions.UpdateReference .NET example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, accesses the first worksheet, configures DeleteOptions with UpdateReference set to true, deletes all blank columns on that sheet, and saves the cleaned workbook to output.xlsx, handling any runtime errors.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure delete options to update references after deleting blank columns
            DeleteOptions deleteOptions = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank columns on this worksheet
            worksheet.Cells.DeleteBlankColumns(deleteOptions);

            // Save the modified workbook
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
