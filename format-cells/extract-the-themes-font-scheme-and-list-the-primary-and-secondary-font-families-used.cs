using System;
using Aspose.Cells;

namespace AsposeCellsThemeFontExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Retrieve the primary (major) theme font
            string primaryFont = workbook.Settings.GetThemeFont(FontSchemeType.Major);

            // Retrieve the secondary (minor) theme font
            string secondaryFont = workbook.Settings.GetThemeFont(FontSchemeType.Minor);

            // Output the font families
            Console.WriteLine("Primary (Major) Theme Font: " + primaryFont);
            Console.WriteLine("Secondary (Minor) Theme Font: " + secondaryFont);

            // Optionally save the workbook to demonstrate persistence
            workbook.Save("ThemeFontDemo.xlsx");
        }
    }
}