// Title: Replace an Excel workbook's theme font with a custom font and update all cell styles using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets the workbook's default Font.Name to a custom family, and applies that font to every cell's style before saving. | Show how to loop through all worksheets and used cells in an Aspose.Cells workbook to change each cell's Font.Name while preserving other style attributes. | Explain how to create the output directory if it does not exist when saving a workbook after changing its theme font with Aspose.Cells.
// Common Searches: Aspose.Cells C# change workbook theme font to Calibri and apply to all cells | C# code to set default font for entire Excel file using Aspose.Cells | How to update font for every used cell in an Excel workbook with Aspose.Cells .NET | Save modified workbook after changing default style font with Aspose.Cells
// Tags: set workbook default font Aspose.Cells C# | apply custom font to all cell styles Aspose.Cells | update Excel theme font programmatically .NET | iterate cells to change font Aspose.Cells | save modified workbook with new font Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook), sets a custom default font for the workbook, iterates through all used cells in each worksheet to apply the font to each cell's style, ensures the output directory exists, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";
            string customFont = "Calibri";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook
            }

            // Set the default font for the workbook
            workbook.DefaultStyle.Font.Name = customFont;

            // Update all cells' styles to use the custom font
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells only for efficiency
                foreach (Cell cell in cells)
                {
                    Style style = cell.GetStyle();
                    style.Font.Name = customFont;
                    cell.SetStyle(style);
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
