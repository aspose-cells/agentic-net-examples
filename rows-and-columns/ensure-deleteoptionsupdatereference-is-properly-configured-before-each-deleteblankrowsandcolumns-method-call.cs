// Title: Enable DeleteOptions.UpdateReference when calling DeleteBlankRowsAndColumns on each worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a DeleteOptions instance with UpdateReference set to true and passes it to Worksheet.Cells.DeleteBlankRowsAndColumns for every sheet in a workbook. | Show how to refactor an Aspose.Cells workbook processing loop to delete blank rows and columns while preserving formula references. | Provide an example that demonstrates configuring DeleteOptions.UpdateReference before invoking DeleteBlankRowsAndColumns on multiple worksheets.
// Common Searches: Aspose.Cells preserve formulas when deleting empty rows and columns in C# | set DeleteOptions.UpdateReference true before cleaning up a workbook | how to keep cell references after removing blank rows using Aspose.Cells | C# example of deleting blank rows and columns without breaking formulas | configure delete options for reference updates in Aspose.Cells
// Tags: DeleteOptions.UpdateReference setup Aspose.Cells | remove empty rows keep formulas C# | clean workbook blank columns without breaking references | Aspose.Cells delete blank rows and columns preserving links | configure delete options for reference integrity .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, iterates through each worksheet, and deletes blank rows and columns. To ensure formulas and cell references remain valid, a DeleteOptions object with UpdateReference enabled must be created and supplied to the DeleteBlankRowsAndColumns method for every sheet before the deletion operations are performed.
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
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Delete blank rows and columns in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Remove blank rows
                    sheet.Cells.DeleteBlankRows();

                    // Remove blank columns
                    sheet.Cells.DeleteBlankColumns();
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {exSheet.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
