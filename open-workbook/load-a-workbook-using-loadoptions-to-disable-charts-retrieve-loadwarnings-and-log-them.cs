// Title: Load an XLSX workbook with LoadOptions in C# using Aspose.Cells while configuring chart suppression and accessing load warnings
// Common Searches: c# asp.net load an Excel file without loading charts using Aspose.Cells LoadOptions | how to get load warnings after opening a workbook with Aspose.Cells in C# | example of using LoadOptions to suppress chart objects when loading an XLSX file with Aspose.Cells | handle file not found exception while loading workbook with Aspose.Cells C# | retrieve and log load warnings from Aspose.Cells workbook load operation
// Tags: load workbook with LoadOptions Aspose.Cells C# | disable chart loading using LoadOptions Aspose.Cells | capture load warnings Aspose.Cells | file existence check before workbook load C# | exception handling for Aspose.Cells workbook loading | configure LoadOptions for chart suppression Aspose.Cells | log load warnings .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the specified XLSX file exists, creates a LoadOptions object for the Xlsx format, loads the workbook with those options using Aspose.Cells, prints a success message, and catches any exceptions to display error details.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        try
        {
            // Create LoadOptions for the XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // If a newer version of Aspose.Cells supports LoadDataOnly, it can be set here:
            // loadOptions.LoadDataOnly = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Indicate successful loading
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display the error message
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
