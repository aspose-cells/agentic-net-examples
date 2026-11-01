// Title: Apply a consistent column width to all worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, sets Cells.StandardWidth to 15 for every worksheet, and saves the result to a new file. | Show how to loop through workbook.Worksheets and assign a uniform column width using the Cells.StandardWidth property in Aspose.Cells. | Create a console application that validates the input file, applies a standard column width across all sheets, and handles any exceptions.
// Common Searches: Aspose.Cells C# set same column width for all sheets in a workbook | How to use Cells.StandardWidth to standardize column width across multiple worksheets | Batch update column width in Excel file with Aspose.Cells .NET | C# example for applying default column width to every worksheet using Aspose.Cells
// Tags: Aspose.Cells set Cells.StandardWidth | uniform column width across worksheets | batch column width adjustment .NET | Excel workbook column sizing with Aspose | C# apply standard column width

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (input.xlsx), iterates through each worksheet, sets the default column width to 15 characters via Cells.StandardWidth, and saves the modified file as output.xlsx, including input‑file validation and exception handling.
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

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Desired standard column width (in characters)
            double standardWidth = 15.0;

            // Apply the standard width to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use Cells.StandardWidth to set the default column width
                sheet.Cells.StandardWidth = standardWidth;
            }

            // Save the updated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
