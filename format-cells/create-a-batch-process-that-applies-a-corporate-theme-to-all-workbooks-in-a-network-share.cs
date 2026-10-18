// Title: Batch apply a corporate .thmx theme to Excel workbooks on a network share using Aspose.Cells for .NET
// Common Searches: how to apply a .thmx theme to multiple Excel files with Aspose.Cells in C# | C# code to recursively apply a theme to all workbooks in a shared folder | using reflection to call ApplyTheme in Aspose.Cells when the method is not directly exposed | determine correct SaveFormat for different Excel extensions in Aspose.Cells | batch processing Excel files on a network share with Aspose.Cells
// Tags: corporate theme application Aspose.Cells | batch Excel workbook processing network share | reflection based ApplyTheme invocation | extension based SaveFormat selection Aspose.Cells | recursive Excel file discovery .NET

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace CorporateThemeBatch
{
    // The example scans a network share for every Excel workbook, loads each file with Aspose.Cells, uses reflection to invoke the ApplyTheme method with a corporate .thmx theme, selects the appropriate SaveFormat based on the original file extension, overwrites the original workbook, and logs progress and errors for robust batch processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Network share containing the workbooks
            string workbooksFolder = @"\\Server\Share\Workbooks";

            // Path to the corporate theme file (.thmx)
            string corporateThemePath = @"\\Server\Share\Themes\CorporateTheme.thmx";

            // Verify that the theme file exists
            if (!File.Exists(corporateThemePath))
            {
                Console.WriteLine($"Corporate theme file not found: {corporateThemePath}");
                return;
            }

            // Get all Excel files recursively from the network share
            string[] excelFiles = Directory.GetFiles(workbooksFolder, "*.xls*", SearchOption.AllDirectories);

            foreach (string filePath in excelFiles)
            {
                try
                {
                    // Ensure the workbook file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Workbook file not found (skipped): {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Apply the corporate theme using reflection (compatible with multiple Aspose.Cells versions)
                    MethodInfo applyThemeMethod = typeof(Workbook).GetMethod("ApplyTheme", new[] { typeof(string) });
                    if (applyThemeMethod != null)
                    {
                        applyThemeMethod.Invoke(workbook, new object[] { corporateThemePath });
                    }
                    else
                    {
                        Console.WriteLine("ApplyTheme method not available in this Aspose.Cells version; theme not applied.");
                    }

                    // Determine the appropriate save format based on original extension
                    SaveFormat format = GetSaveFormat(Path.GetExtension(filePath));

                    // Save the workbook, overwriting the original file
                    workbook.Save(filePath, format);

                    Console.WriteLine($"Theme applied and saved: {filePath}");
                }
                catch (Exception ex)
                {
                    // Log any errors but continue processing other files
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }

        // Helper method to map file extensions to Aspose.Cells SaveFormat values
        private static SaveFormat GetSaveFormat(string extension)
        {
            switch (extension.ToLower())
            {
                case ".xlsx":
                    return SaveFormat.Xlsx;
                case ".xls":
                    return SaveFormat.Excel97To2003;
                case ".xlsm":
                    return SaveFormat.Xlsm;
                case ".xlsb":
                    return SaveFormat.Xlsb;
                case ".csv":
                    return SaveFormat.Csv;
                default:
                    // Default to XLSX if unknown
                    return SaveFormat.Xlsx;
            }
        }
    }
}
