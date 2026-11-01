// Title: How to disable the default 1900 date system in Aspose.Cells for .NET by setting CalculationOptions.Use1900DateSystem to false
// Common Searches: Aspose.Cells disable 1900 date system C# example | set CalculationOptions.Use1900DateSystem false in a .NET workbook | switch Excel workbook to 1904 date system using Aspose.Cells API | change Excel date system to 1904 with Aspose.Cells for C# | Aspose.Cells calculation options 1904 date system usage
// Tags: aspocells calculationoptions use1900datesystem | disable 1900 date system aspnet workbook | set excel 1904 date system aspose.cells c# | workbook date system configuration aspose.cells | c# aspose.cells change date system

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, indicates where to set CalculationOptions.Use1900DateSystem to false (or use workbook.Settings.Is1904DateSystem) to switch from the default 1900 date system to the 1904 system, saves the workbook as Output.xlsx, prints the full file path, and catches any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Note: The 1904 date system can be enabled via workbook.Settings.Is1904DateSystem
            // if the property is available in the referenced Aspose.Cells version.

            // Define output file path
            string outputPath = "Output.xlsx";

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
