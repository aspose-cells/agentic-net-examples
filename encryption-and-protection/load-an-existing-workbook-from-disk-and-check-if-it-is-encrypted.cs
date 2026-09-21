// Title: Determine whether a local .xlsx workbook is password‑protected using Aspose.Cells for .NET
// AI Prompts: Write C# code that attempts to open an .xlsx file with Aspose.Cells and indicates if the file is encrypted by catching the relevant CellsException. | Demonstrate how to configure LoadOptions and handle a password‑required exception to report the encryption status of a workbook without supplying a password.
// Common Searches: asp.net check if excel file is encrypted before opening with Aspose.Cells | c# detect password protected workbook using Aspose.Cells LoadOptions | how to identify encrypted .xlsx using Aspose.Cells exception handling | determine if Excel workbook requires password in C# Aspose.Cells
// Tags: detect encrypted Excel workbook Aspose.Cells | LoadOptions password detection C# | CellsException handling for protected workbook | verify workbook encryption status .NET | Aspose.Cells workbook protection check

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a workbook from a specified path with Aspose.Cells. If loading succeeds, the file is not encrypted; if a CellsException containing a password‑related message is thrown, the code reports the workbook as password‑protected.
class Program
{
    static void Main()
    {
        // Path to the workbook file on disk
        string filePath = @"C:\Path\To\Your\Workbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook (throws if the file is password‑protected)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Workbook loaded successfully – it is not encrypted
            Console.WriteLine("The workbook is not encrypted.");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
        }
        catch (CellsException ex)
        {
            // Check if the exception indicates a password‑protected workbook
            if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is encrypted.");
                // Optional: prompt for password and reload with LoadOptions.Password
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle other unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
