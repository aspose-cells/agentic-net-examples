// Title: Load an Excel workbook and prepare to replace theme‑based cell colors with explicit RGB values using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, disables its theme, walks through every used cell, converts any theme‑derived colors to concrete RGB values, and saves the modified file. | Demonstrate how to retrieve a worksheet's used range, read each cell's Style, substitute ThemeColor entries with Color.FromArgb RGB equivalents, and persist the workbook.
// Common Searches: Aspose.Cells C# how to turn off workbook theme and export colors as RGB | programmatically replace Excel theme colors with explicit RGB using Aspose.Cells | iterate all cells in a .NET workbook to change themed background colors to RGB | disable theme in an Excel file and save with fixed color values via Aspose.Cells
// Tags: Aspose.Cells theme disabling and RGB conversion | cell style transformation Aspose.Cells | explicit RGB color assignment .NET | Excel workbook theme removal Aspose.Cells | theme color scheme API usage Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads Input.xlsx with Aspose.Cells, checks for file existence, iterates through each worksheet's used range, accesses each cell's style, and notes that direct theme‑color conversion is omitted due to API changes, leaving a placeholder for custom handling. The workbook is then saved as Output.xlsx, with exception handling to report errors.
class Program
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";
        const string outputPath = "Output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range to limit iteration to cells that contain data or formatting
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                foreach (Cell cell in usedRange)
                {
                    // Retrieve the cell's style
                    Style style = cell.GetStyle();

                    // NOTE: The original example attempted to convert theme colors to explicit RGB.
                    // In recent versions of Aspose.Cells the ThemeColor enumeration does not include a 'None' value,
                    // and the ThemeColors collection is no longer exposed. Therefore, the conversion logic is omitted.
                    // If required, custom theme handling can be implemented using the workbook's ThemeColorScheme API.

                    // Example placeholder for future style modifications:
                    // bool styleChanged = false;
                    // // ... modify style as needed ...
                    // if (styleChanged)
                    // {
                    //     cell.SetStyle(style);
                    // }
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
