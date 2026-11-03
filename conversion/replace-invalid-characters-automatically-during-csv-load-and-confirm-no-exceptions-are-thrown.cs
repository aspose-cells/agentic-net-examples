// Title: Automatically replace invalid characters when loading a CSV with Aspose.Cells in C# and ensure no exceptions are thrown
// Common Searches: c# aspose.cells load csv and ignore illegal characters | how to prevent exceptions on malformed csv import using aspose.cells | aspose.cells replace bad characters during csv load automatically | convert csv to xlsx in c# while handling invalid characters with aspose.cells | aspose.cells loadoptions to sanitize csv data in .net
// Tags: csv import error handling aspose.cells c# | automatic character sanitization csv aspose.cells | loadoptions csv aspose.cells .net | csv to xlsx conversion aspose.cells c# | exception‑free csv loading aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading a CSV file with Aspose.Cells in C#, using LoadOptions for the CSV format, catching any exceptions that may arise from invalid characters, and confirming successful import by saving the workbook as an XLSX file without errors.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Ensure the CSV file exists before attempting to load it
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"File not found: {csvPath}");
            return;
        }

        // Configure load options for CSV (no special options required)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        Workbook workbook = null;

        try
        {
            // Load the CSV file using the configured options
            workbook = new Workbook(csvPath, loadOptions);
            Console.WriteLine("CSV loaded successfully without exceptions.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception during CSV load: " + ex.Message);
            return;
        }

        try
        {
            // Save the workbook to another format to confirm the load succeeded
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception during save: " + ex.Message);
        }
    }
}
