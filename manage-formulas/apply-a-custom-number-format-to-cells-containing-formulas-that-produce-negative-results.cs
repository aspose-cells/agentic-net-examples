// Title: Apply a red custom number format to formula cells that evaluate to negative values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, calculates all formulas, scans the used range, and assigns the custom format "#,##0.00;[Red]-#,##0.00" only to formula cells whose result is a negative number. | Show how to retrieve a cell’s evaluated value, verify it is a negative double, and programmatically apply a custom style with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells C# apply red number format only to negative results of formula cells | how to set custom number format for negative values after formula calculation in .NET | iterate over used cells in Excel workbook with Aspose.Cells and format negatives | C# example for conditional number formatting of formula results using Aspose.Cells | apply custom number format to negative numbers without affecting positive cells in Aspose.Cells
// Tags: custom number format negative values Aspose.Cells | format formula cells based on evaluated result C# | iterate used range set cell style Aspose.Cells | calculate all formulas before applying formatting Aspose | red negative number format Excel C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, forces calculation of all formulas, iterates through the used cells, and applies the custom red number format "#,##0.00;[Red]-#,##0.00" to any formula cell whose evaluated value is a negative double, then saves the updated file.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string customFormat = "#,##0.00;[Red]-#,##0.00";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure all formulas are calculated
            workbook.CalculateFormula();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate through used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Retrieve the evaluated value of the formula
                        object value = cell.Value;

                        // Apply custom format if the value is a negative number
                        if (value is double numericValue && numericValue < 0)
                        {
                            Style style = cell.GetStyle();
                            style.Custom = customFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
