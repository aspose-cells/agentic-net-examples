// Title: Apply a Gray25 small‑dot fill pattern to cells containing the word “placeholder” with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing XLSX file, creates a Gray25 pattern style with black dots on a white background, applies it to every cell whose text includes the word placeholder, and saves the workbook. | Show how to iterate over the used range of a worksheet in Aspose.Cells, detect cells containing a specific substring, and set a custom fill pattern on those cells. | Adapt the example to use a different dot pattern such as LightDown while still targeting cells that contain placeholder text.
// Common Searches: Aspose.Cells C# set Gray25 fill pattern for cells with specific text | how to apply background pattern to placeholder cells in Excel using Aspose.Cells | C# iterate over used cells and change style based on cell value Aspose | replace placeholder text cells with dot pattern in .NET Excel library | save workbook after applying custom fill style with Aspose.Cells
// Tags: gray25 dot fill Aspose.Cells | style cells containing placeholder C# | apply background pattern Excel .NET | iterate used range Aspose.Cells | save workbook after styling Aspose

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The code loads input.xlsx, creates a Gray25 style (black dots on white), scans all used cells for the word “placeholder”, applies the style to matching cells, and saves the result as output.xlsx.
class ApplyDotPatternToPlaceholderCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style with a small dots fill pattern
            Style dotStyle = workbook.CreateStyle();
            // Use a pattern that exists in the current Aspose.Cells version (e.g., Gray25)
            dotStyle.Pattern = BackgroundType.Gray25;          // Approximate small dots pattern
            dotStyle.ForegroundColor = Color.Black;           // Dot color
            dotStyle.BackgroundColor = Color.White;           // Background color

            // Iterate through all used cells and apply the style to placeholders
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell?.Value != null && cell.Value.ToString().Contains("placeholder"))
                    {
                        cell.SetStyle(dotStyle);
                    }
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
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
