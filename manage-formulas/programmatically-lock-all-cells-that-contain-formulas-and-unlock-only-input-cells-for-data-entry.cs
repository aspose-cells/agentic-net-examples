// Title: Lock formula cells and unlock data entry cells in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells lock cells that contain formulas while leaving other cells editable | How to protect an Excel worksheet without a password using Aspose.Cells in .NET | Set IsLocked property for formula cells only with Aspose.Cells C# example | Unlock user input cells and lock calculated cells in a workbook using Aspose.Cells | Iterate through all cells in Aspose.Cells to apply protection based on formula presence
// Tags: lock formula cells Aspose.Cells .NET | unlock data entry cells Aspose.Cells | worksheet protection without password Aspose.Cells | set cell IsLocked property C# | iterate cells to apply protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, loops through each worksheet and every used cell, sets the IsLocked style flag to true for cells containing formulas and false for all other cells, reapplies the style, protects each sheet without a password, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Retrieve the current style
                    Style style = cell.GetStyle();

                    // Lock formula cells, unlock others
                    if (cell.IsFormula)
                    {
                        style.IsLocked = true;   // Lock formula cells
                    }
                    else
                    {
                        style.IsLocked = false;  // Unlock input cells
                    }

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }

                // Protect the worksheet (no password)
                // The overload requires both password and oldPassword parameters
                sheet.Protect(ProtectionType.All, null, null);
            }

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
