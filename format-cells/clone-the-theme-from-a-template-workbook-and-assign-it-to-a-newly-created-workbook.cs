using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeCloneDemo
{
    public class ThemeCloneExample
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string templatePath = "template.xlsx";
                const string outputPath = "output_with_theme.xlsx";

                // Verify that the template file exists
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Load the template workbook that contains the desired theme
                Workbook templateWorkbook = new Workbook(templatePath);

                // Create a new empty workbook
                Workbook newWorkbook = new Workbook();

                // Copy the theme from the template workbook to the new workbook
                newWorkbook.CopyTheme(templateWorkbook);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the new workbook with the cloned theme
                newWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}