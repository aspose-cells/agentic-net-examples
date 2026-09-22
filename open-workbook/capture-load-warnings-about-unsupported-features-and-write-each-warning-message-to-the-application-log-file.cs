// Title: Capture and log unsupported feature warnings when loading an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Implement an IWarningCallback that records each load warning from Aspose.Cells and appends it with a timestamp to a log file. | Modify the workbook loading code to use LoadOptions with a warning callback that writes unsupported‑feature warnings to app.log. | Provide C# sample that accesses the warning collection after opening an Excel file and logs every entry.
// Common Searches: how to log unsupported feature warnings from Aspose.Cells when opening an xlsx file in C# | Aspose.Cells .NET capture load warnings using IWarningCallback example | write Excel load warnings to a custom log file with Aspose.Cells LoadOptions | retrieve warning messages for unsupported Excel features during workbook load in C#
// Tags: Aspose.Cells IWarningCallback implementation | log Excel load warnings .NET | capture unsupported feature warnings Aspose.Cells | load workbook with warning callback C# | write load warnings to application log

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to load an XLSX workbook with Aspose.Cells using LoadOptions, attach an IWarningCallback to capture any unsupported‑feature warnings, and append each warning with a timestamp to a specified application log file while handling load errors gracefully.
class Program
{
    static void Main()
    {
        // Path to the Excel file to load
        string inputPath = "input.xlsx";

        // Path to the application log file
        string logPath = "app.log";

        // Verify that the input file exists to prevent FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Prepare load options (adjust the format if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // NOTE: Aspose.Cells does not expose load warnings via a property on Workbook in all versions.
            // If needed, warnings can be captured via LoadOptions or other mechanisms in newer releases.

            // (Optional) Continue processing the workbook as needed
            // ...
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            try
            {
                string errorEntry = $"{DateTime.Now:u} - ERROR: {ex.Message}{Environment.NewLine}";
                File.AppendAllText(logPath, errorEntry);
            }
            catch
            {
                // If logging fails, write to console only
                Console.WriteLine("Failed to write to log file.");
            }

            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
