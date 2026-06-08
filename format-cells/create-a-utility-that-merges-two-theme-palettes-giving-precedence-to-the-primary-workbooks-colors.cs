using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeMergeApp
{
    public static class ThemeMerger
    {
        // Merges the theme palettes of two workbooks. Colors from the primary workbook take precedence.
        public static void MergeThemes(string primaryPath, string secondaryPath, string outputPath)
        {
            // Validate input files
            if (!File.Exists(primaryPath))
                throw new FileNotFoundException($"Primary workbook not found: {primaryPath}");
            if (!File.Exists(secondaryPath))
                throw new FileNotFoundException($"Secondary workbook not found: {secondaryPath}");

            // Load the primary workbook (the one that will receive the merged theme)
            Workbook primaryWorkbook = new Workbook(primaryPath);

            // Load the secondary workbook (the source of additional theme colors)
            Workbook secondaryWorkbook = new Workbook(secondaryPath);

            // Iterate through all theme color types and merge
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                Color primaryColor = primaryWorkbook.GetThemeColor(type);
                Color secondaryColor = secondaryWorkbook.GetThemeColor(type);
                Color mergedColor = primaryColor != Color.Empty ? primaryColor : secondaryColor;
                primaryWorkbook.SetThemeColor(type, mergedColor);
            }

            // Save the workbook with the merged theme palette
            primaryWorkbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // Expect three arguments: primaryPath secondaryPath outputPath
                if (args.Length != 3)
                {
                    Console.WriteLine("Usage: ThemeMergeApp <primary.xlsx> <secondary.xlsx> <output.xlsx>");
                    return;
                }

                string primaryPath = args[0];
                string secondaryPath = args[1];
                string outputPath = args[2];

                ThemeMerger.MergeThemes(primaryPath, secondaryPath, outputPath);
                Console.WriteLine($"Merged workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}