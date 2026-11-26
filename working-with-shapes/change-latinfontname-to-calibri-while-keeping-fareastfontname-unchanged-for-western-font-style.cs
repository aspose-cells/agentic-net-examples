// Title: Set the Latin (Western) font to Calibri for all cells in an Excel workbook using Aspose.Cells for .NET while preserving FarEast font settings
// Common Searches: Aspose.Cells C# change cell font to Calibri without affecting FarEast font | how to update Latin font for every used cell in a workbook with Aspose.Cells .NET | preserve FarEastFontName while setting Western font in Excel using Aspose.Cells | apply Calibri style to all cells in an Excel file with Aspose.Cells C# | iterate worksheets and set Font.Name to Calibri using Aspose.Cells
// Tags: latin font name Calibri Aspose.Cells | retain FarEastFontName Aspose.Cells | apply cell style to entire workbook Aspose.Cells | bulk font change across worksheets C# | modify western font without affecting FarEast font Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, loops through each worksheet and every used cell, changes the cell style's Latin (Western) font to "Calibri" while leaving the FarEastFontName unchanged, and saves the updated workbook to a new file.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                var usedRange = sheet.Cells.MaxDisplayRange; // Aspose.Cells.Range

                // Iterate over each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Retrieve the current style of the cell
                    Style style = cell.GetStyle();

                    // Set the Western (Latin) font to Calibri
                    style.Font.Name = "Calibri";

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
