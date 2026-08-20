// Title: Get Workbook Theme Font Families (Major & Minor) with Aspose.Cells for .NET
// Description: Shows how to read the major (primary) and minor (secondary) font families defined by an Excel workbook’s theme using Aspose.Cells for .NET (Workbook.Settings.GetThemeFont). The sample prints the fonts to the console and saves the workbook.
// Keywords: Aspose.Cells GetThemeFont | C# read Excel theme fonts | major theme font | minor theme font | Excel workbook theme font families | FontSchemeType | Aspose.Cells Settings | extract theme fonts .NET
// Common Searches: Aspose.Cells get major theme font C# | How to read minor theme font from Excel using Aspose.Cells | Retrieve default theme fonts in a new workbook .NET | C# code to list Excel theme font families with Aspose
// Developer Intent: Retrieve the major and minor font names defined by a workbook’s theme.
// Use Cases: Display default theme fonts to users before applying custom styles. | Log theme font names for compliance with corporate branding. | Validate that an imported workbook matches expected theme typography.
// AI Prompts: Generate C# code with Aspose.Cells to change the major theme font to 'Calibri' and the minor theme font to 'Arial'. | Create a method that returns a (string majorFont, string minorFont) tuple from a given Workbook object. | Add robust error handling for GetThemeFont when loading a workbook from a stream or corrupted file.

using System;
using Aspose.Cells;

namespace ThemeFontExtractor
{
    // Shows how to read the major (primary) and minor (secondary) font families defined by an Excel workbook’s theme using Aspose.Cells for .NET (Workbook.Settings.GetThemeFont). The sample prints the fonts to the console and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default theme will be applied)
            Workbook workbook = new Workbook();

            // Retrieve the primary (major) and secondary (minor) theme fonts
            string majorFont = workbook.Settings.GetThemeFont(FontSchemeType.Major);
            string minorFont = workbook.Settings.GetThemeFont(FontSchemeType.Minor);

            // Output the font families
            Console.WriteLine("Primary (Major) Theme Font: " + majorFont);
            Console.WriteLine("Secondary (Minor) Theme Font: " + minorFont);

            // Optionally save the workbook if you want to inspect the theme in the file
            workbook.Save("ThemeFontExample.xlsx");
        }
    }
}
