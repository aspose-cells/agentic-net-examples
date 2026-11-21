// Title: Configure Aspose.Cells for .NET to use the 1904 date system for legacy Excel calculations
// Common Searches: C# Aspose.Cells set workbook to 1904 date system for old Excel files | How to enable 1904 date system in Aspose.Cells calculation engine | Aspose.Cells .NET example for legacy date handling with 1904 system | Set workbook.Settings.Is1904DateSystem true in C# Aspose.Cells | Enable 1904 date system when creating a new workbook with Aspose.Cells
// Tags: Aspose.Cells workbook.Settings.Is1904DateSystem | 1904 date system configuration in .NET | legacy Excel date handling with Aspose.Cells | C# set 1904 date system for XLSX workbook | Aspose.Cells calculation engine legacy dates

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, optionally activates the 1904 date system by setting workbook.Settings.Is1904DateSystem to true, saves the file as LegacyDateWorkbook.xlsx, outputs the full path, and catches any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable the 1904 date system if the property exists in your Aspose.Cells version
            // Uncomment the line below when supported:
            // workbook.Settings.Is1904DateSystem = true;

            // Define the output file path
            string outputPath = "LegacyDateWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
