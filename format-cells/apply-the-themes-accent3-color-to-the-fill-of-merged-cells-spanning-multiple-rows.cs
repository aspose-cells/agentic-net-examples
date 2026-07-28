// Title: Apply Theme Accent3 Fill to Merged Cells with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, merges the range A1:B3, defines a solid fill style using the theme's Accent3 color (no tint), applies the style to the merged range, and saves the result as MergedAccent3.xlsx.
// Keywords: Aspose.Cells C# theme color | Accent3 fill merged cells | solid background Aspose.Cells | apply theme color .NET | merged range styling
// Common Searches: Aspose.Cells set Accent3 fill for merged range C# | how to use theme colors in merged cells Aspose | C# apply solid theme background to merged cells | Aspose.Cells merge cells and style with theme color | theme accent fill merged cells .NET
// Developer Intent: Use the workbook's Accent3 theme color as a solid background for a merged cell block.
// Use Cases: Header rows that span multiple columns and need corporate Accent3 shading. | Section titles in templates that automatically adopt the theme's secondary accent. | Dashboard tiles showing key metrics highlighted with the Accent3 background.
// AI Prompts: Generate C# code with Aspose.Cells that merges D5:F7 and fills it with Accent2 at 30% tint. | Show how to apply a solid Accent4 fill and bold font to a merged range in Aspose.Cells. | Write a reusable method that takes a ThemeColorType and a range address, merges the range, and applies the chosen theme color as a solid background.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsThemeAccent3
{
    // Creates a new workbook, merges the range A1:B3, defines a solid fill style using the theme's Accent3 color (no tint), applies the style to the merged range, and saves the result as MergedAccent3.xlsx.
    public class ApplyAccent3ToMergedCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define a range that spans 3 rows and 2 columns (A1:B3)
                AsposeRange mergedRange = worksheet.Cells.CreateRange(0, 0, 3, 2);

                // Merge the cells in the defined range
                mergedRange.Merge();

                // Put a sample value in the merged cell (top‑left cell of the range)
                worksheet.Cells[0, 0].PutValue("Merged Cells with Accent3 Fill");

                // Create a style for the fill
                Style fillStyle = workbook.CreateStyle();

                // Set solid fill pattern
                fillStyle.Pattern = BackgroundType.Solid;

                // Apply the theme's Accent3 color (no tint)
                fillStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0.0);

                // Apply the style to the merged range
                mergedRange.SetStyle(fillStyle);

                // Save the workbook
                workbook.Save("MergedAccent3.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyAccent3ToMergedCells.Run();
        }
    }
}
