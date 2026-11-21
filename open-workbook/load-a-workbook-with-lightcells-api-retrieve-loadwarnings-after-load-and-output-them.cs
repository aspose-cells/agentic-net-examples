// Title: Load an XLSX workbook with Aspose.Cells in C# and handle the absence of load warnings
// Common Searches: how to load an xlsx file with Aspose.Cells in C# and check for load warnings | Aspose.Cells C# example for loading a workbook using LoadOptions | retrieve workbook load diagnostics when LightCells API does not expose warnings | C# code to open an Excel file with Aspose.Cells and handle load errors
// Tags: Aspose.Cells LoadOptions workbook loading C# | Excel XLSX load diagnostics Aspose.Cells | handling workbook load errors C# | load warnings not exposed Aspose.Cells version | LightCells API workbook loading C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an XLSX file, creates LoadOptions for the Xlsx format, and loads the workbook using Aspose.Cells in C#. It demonstrates standard error handling but notes that the current Aspose.Cells release does not provide direct access to load warnings, so only a success message is displayed.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Initialize LoadOptions (LightCells API is not required for this example)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Since the current Aspose.Cells version does not expose load warnings directly,
            // we simply confirm that the workbook was loaded successfully.
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during loading or processing
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
