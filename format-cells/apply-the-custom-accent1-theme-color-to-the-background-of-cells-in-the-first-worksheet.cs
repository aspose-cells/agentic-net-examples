// Title: Set the Accent1 theme color as the background fill for all cells in the first worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply Accent1 theme color to entire worksheet background | How to fill all cells in first sheet with a solid theme color using Aspose.Cells .NET | Apply style flag for cell shading to whole worksheet in Aspose.Cells C# | Set solid fill color of a worksheet using Aspose.Cells style and StyleFlag | Create workbook and set Accent1 background for first sheet with Aspose.Cells
// Tags: accent1 theme background style Aspose.Cells | cell shading styleflag Aspose.Cells | solid fill color worksheet Aspose.Cells | apply style to all cells Aspose.Cells | theme color formatting Aspose.Cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, accesses the first worksheet, defines a style with a solid fill using the Accent1 theme RGB value, enables only cell shading via a StyleFlag, applies the style to every cell in the sheet, and saves the file as FirstWorksheet_Accent1Background.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a style with a solid fill using a color similar to the Accent1 theme color
                Style accentStyle = workbook.CreateStyle();
                accentStyle.Pattern = BackgroundType.Solid;                 // Solid fill
                accentStyle.ForegroundColor = Color.FromArgb(0, 112, 192);   // Approximate Accent1 color
                // No tint needed for solid fill

                // Apply the style to all cells in the first worksheet
                // The StyleFlag ensures only the cell shading (background) is affected
                StyleFlag flag = new StyleFlag
                {
                    CellShading = true
                };
                worksheet.Cells.ApplyStyle(accentStyle, flag);

                // Define output file path
                string outputPath = "FirstWorksheet_Accent1Background.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
