// Title: Set header row font color to the workbook’s Accent1 theme color across all worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change first row font color to workbook theme Accent1 for every sheet | Iterate through all worksheets and apply theme accent color to header cells in .NET | GetThemeColor Accent1 usage to format header row in Excel with Aspose.Cells | Set font color of header row based on workbook theme in C# Aspose.Cells example | Apply workbook theme color to header row across multiple worksheets programmatically
// Tags: set header row font color using GetThemeColor | use workbook theme Accent1 for header formatting | iterate worksheets Aspose.Cells C# | format first row cells Aspose.Cells | theme-based font styling in .NET Excel

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook (or creates a new one if missing), retrieves the workbook’s Accent1 theme color via GetThemeColor, loops through each worksheet, and updates the font color of every cell in the first row to that theme color before saving the file, with robust error handling for loading, processing, and saving.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex)
                {
                    // If the workbook is password protected or any other loading issue occurs,
                    // fall back to creating a new workbook.
                    Console.WriteLine($"Failed to load workbook \"{inputPath}\": {ex.Message}");
                    workbook = new Workbook();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error loading workbook \"{inputPath}\": {ex.Message}");
                    workbook = new Workbook();
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Ensure workbook is instantiated.
            if (workbook == null)
            {
                Console.WriteLine("Unable to create or load a workbook. Exiting.");
                return;
            }

            // Retrieve the theme's Accent1 color.
            Color accent1 = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Apply the Accent1 color to header cells of each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    int lastColumn = sheet.Cells.MaxColumn;

                    for (int col = 0; col <= lastColumn; col++)
                    {
                        // Get current style, modify font color, and reapply.
                        Style style = sheet.Cells[0, col].GetStyle();
                        style.Font.Color = accent1;
                        sheet.Cells[0, col].SetStyle(style);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the modified workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook to \"{outputPath}\": {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
