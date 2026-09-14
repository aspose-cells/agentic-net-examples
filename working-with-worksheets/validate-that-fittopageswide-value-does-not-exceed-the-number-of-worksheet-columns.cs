// Title: Validate and Correct Worksheet FitToPagesWide Setting According to Column Count with Aspose.Cells for .NET
// AI Prompts: Write C# code that loops through every worksheet in a workbook, compares PageSetup.FitToPagesWide to the actual number of columns (Cells.MaxColumn+1), logs worksheets where the value is too high, and resets it to the column count. | Generate a C# example that loads an Excel file using Aspose.Cells, checks if FitToPagesWide exceeds the worksheet's column limit, updates the setting to a valid value, and saves the modified file.
// Common Searches: Aspose.Cells C# how to ensure FitToPagesWide does not exceed worksheet columns | programmatically adjust FitToPagesWide based on MaxColumn in a .NET workbook | check page setup FitToPagesWide against column count before saving Excel with Aspose | C# validate FitToPagesWide setting for each sheet in an Aspose.Cells workbook
// Tags: validate FitToPagesWide column limit Aspose.Cells | adjust worksheet page setup FitToPagesWide C# | compare FitToPagesWide to Cells.MaxColumn Aspose | Aspose.Cells page setup overflow prevention | C# workbook column count validation

using Aspose.Cells;
using System;

// Loads a workbook, iterates each worksheet, compares PageSetup.FitToPagesWide with the actual column count (Cells.MaxColumn+1), logs any excess, resets FitToPagesWide to the column count when necessary, and saves the updated workbook.
class FitToPagesWideValidator
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the FitToPagesWide setting from the worksheet's PageSetup
            int fitToPagesWide = sheet.PageSetup.FitToPagesWide;

            // Determine the total number of columns in the worksheet.
            // Excel supports up to 16384 columns (A to XFD). 
            // If you want to consider only columns that contain data, use MaxColumn + 1.
            int totalColumns = sheet.Cells.MaxColumn + 1; // zero‑based index + 1

            // Validate the FitToPagesWide value
            if (fitToPagesWide > totalColumns)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": FitToPagesWide ({fitToPagesWide}) exceeds the number of columns ({totalColumns}).");
                // Optionally, adjust the value to a valid range
                sheet.PageSetup.FitToPagesWide = totalColumns;
                Console.WriteLine($"Adjusted FitToPagesWide to {totalColumns}.");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": FitToPagesWide ({fitToPagesWide}) is within the column limit ({totalColumns}).");
            }
        }

        // Save the workbook if any adjustments were made (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
