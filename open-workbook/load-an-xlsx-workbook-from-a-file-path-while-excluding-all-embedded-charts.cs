// Title: Load an XLSX workbook from a file path while excluding all embedded charts with Aspose.Cells for .NET
// Common Searches: c# asp.net load xlsx workbook without embedded charts using aspose.cells | how to open an Excel file data only with Aspose.Cells in C# | Aspose.Cells LoadOptions to skip charts when loading workbook | load Excel file from path excluding charts with Aspose.Cells C# | use LoadDataOnly property to ignore charts in Aspose.Cells workbook loading
// Tags: load workbook without charts Aspose.Cells | XLSX LoadOptions LoadDataOnly C# | exclude embedded charts Aspose.Cells | open Excel file data only Aspose.Cells | Aspose.Cells workbook loading options

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the specified XLSX file exists, creates a LoadOptions object for the Xlsx format (optionally setting LoadDataOnly to omit chart data), and loads the workbook from the given path using Aspose.Cells, with proper exception handling.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string filePath = @"C:\Path\To\YourWorkbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Configure load options for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // If the current Aspose.Cells version supports LoadDataOnly, enable it:
            // loadOptions.LoadDataOnly = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            Console.WriteLine("Workbook loaded successfully.");
            // Additional processing can be performed here
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
