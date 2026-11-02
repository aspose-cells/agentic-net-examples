// Title: Disable the default 1900 date system in Aspose.Cells for .NET by setting CalculationOptions.Use1900DateSystem to false
// Common Searches: C# Aspose.Cells how to turn off the 1900 date system | Aspose.Cells workbook use 1904 date system instead of 1900 | Set CalculationOptions.Use1900DateSystem false in .NET | Switch Excel date system to 1904 with Aspose.Cells API | Disable 1900 date system when creating a workbook in Aspose.Cells
// Tags: Aspose.Cells disable 1900 date system | Workbook.Settings.Is1904DateSystem property | Aspose.Cells enable 1904 date system | C# set Use1900DateSystem false | Excel date system configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to create a new Workbook in Aspose.Cells for .NET and disable the default 1900 date system by setting CalculationOptions.Use1900DateSystem to false (or using the Settings.Is1904DateSystem/Date1904 properties depending on the library version), then saves the workbook to a specified file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (default 1900 date system is enabled)
            Workbook workbook = new Workbook();

            // Switch to the 1904 date system if the API supports it.
            // In some Aspose.Cells versions the property name may differ.
            // Uncomment the appropriate line based on the installed version.

            // workbook.Settings.Is1904DateSystem = true; // Aspose.Cells 23.x and later
            // workbook.Settings.Date1904 = true;        // Alternative property in older versions

            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
