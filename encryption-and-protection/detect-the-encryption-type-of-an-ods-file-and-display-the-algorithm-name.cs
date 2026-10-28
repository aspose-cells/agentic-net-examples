// Title: Determine if an ODS workbook is password‑protected with Aspose.Cells for .NET
// Common Searches: how to programmatically check if an ODS file is encrypted using Aspose.Cells in C# | C# detect password protection on OpenDocument Spreadsheet with Aspose.Cells | using LoadOptions to identify encrypted ODS workbooks in .NET | Aspose.Cells catch CellsException for encrypted ODS files example | verify ODS file encryption status before opening with Aspose.Cells
// Tags: Aspose.Cells ODS encryption detection | LoadOptions ODS format .NET | CellsException password protected ODS | C# check ODS workbook encryption | detect encrypted OpenDocument spreadsheet Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS file with Aspose.Cells using LoadOptions for the ODS format. It attempts to open the workbook without a password; if loading succeeds the file is reported as not encrypted, while a CellsException containing a password‑related message indicates the ODS workbook is password‑protected. Other exceptions are logged as general errors.
class Program
{
    static void Main()
    {
        // Path to the ODS file
        string filePath = "sample.ods";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Prepare load options for ODS format (use the non‑obsolete enum value)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods);

        try
        {
            // Attempt to load the workbook without a password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the file is not encrypted
            Console.WriteLine("The ODS file is not encrypted.");
        }
        catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            // Specific handling for password‑protected files
            Console.WriteLine("The ODS file is encrypted.");
        }
        catch (Exception ex)
        {
            // General error handling
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
