using System;
using Aspose.Cells;

namespace ThemeFontExtractor
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Retrieve the primary (major) theme font
            string majorFont = workbook.Settings.GetThemeFont(FontSchemeType.Major);
            // Retrieve the secondary (minor) theme font
            string minorFont = workbook.Settings.GetThemeFont(FontSchemeType.Minor);

            // Display the results
            Console.WriteLine("Primary (Major) Theme Font: " + majorFont);
            Console.WriteLine("Secondary (Minor) Theme Font: " + minorFont);

            // Optionally save the workbook (demonstrates lifecycle compliance)
            workbook.Save("ThemeFontInfo.xlsx");
        }
    }
}