// Title: How to apply the workbook Accent1 theme color as a solid background to every used cell in the first worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Accent1 theme color as cell background for entire sheet | Apply solid theme color to all used cells in first worksheet with Aspose.Cells | GetThemeColor Accent1 and use StyleFlag CellShading in Aspose.Cells .NET | Create a style with workbook theme color and apply to whole worksheet using Aspose.Cells | Set background color of all cells to workbook Accent1 theme in C# Aspose.Cells
// Tags: Workbook.GetThemeColor Accent1 usage | StyleFlag CellShading whole worksheet | Aspose.Cells solid background style C# | apply theme color to all cells Aspose.Cells | create style with theme color Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, retrieves the Accent1 theme color via Workbook.GetThemeColor, builds a solid foreground style, configures a StyleFlag to affect only cell shading, applies the style to every used cell in the first worksheet, and saves the result as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the Accent1 theme color using the workbook API
                Color accentColor = workbook.GetThemeColor(ThemeColorType.Accent1);

                // Create a style that uses the custom Accent1 theme color
                Style accentStyle = workbook.CreateStyle();
                accentStyle.ForegroundColor = accentColor;
                accentStyle.Pattern = BackgroundType.Solid;

                // Apply the style to all used cells in the first worksheet
                StyleFlag flag = new StyleFlag
                {
                    CellShading = true // Apply background color only
                };
                sheet.Cells.ApplyStyle(accentStyle, flag);

                // Define output file path
                string outputPath = "Result.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
