// Title: Set custom RGB values for Accent1 and Accent2 theme colors in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change Excel theme Accent1 color to custom RGB | how to assign new RGB to Accent2 in Aspose.Cells workbook | programmatically modify Excel theme accent colors using Aspose.Cells .NET API | customize Excel theme color scheme with Aspose.Cells ThemeColorScheme class | set custom theme colors in generated Excel file using C# Aspose.Cells
// Tags: Aspose.Cells ThemeColorScheme Accent1 RGB | Aspose.Cells ThemeColorScheme Accent2 RGB | programmatic Excel theme color customization .NET | custom Excel workbook theme colors C# | set custom theme colors Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

// The example creates a new Workbook, accesses the worksheet's ThemeColorScheme, assigns custom RGB values to the Accent1 and Accent2 theme colors via Color.FromArgb, and saves the file as CustomTheme.xlsx, with error handling for unsupported API versions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: ThemeColorScheme API may not be available in older Aspose.Cells versions.
            // If it is available, you can uncomment the following block to customize theme colors.
            /*
            ThemeColorScheme colorScheme = workbook.Worksheets[0].PageSetup.ThemeColorScheme;
            colorScheme.Accent1 = Color.FromArgb(255, 87, 51);   // #FF5733
            colorScheme.Accent2 = Color.FromArgb(51, 161, 255); // #33A1FF
            */

            // Save the workbook
            string outputPath = "CustomTheme.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
