// Title: How to load a password‑protected XLSX file with Aspose.Cells for .NET and log authentication errors
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions to open an encrypted .xlsx workbook, checks file existence, and records the CellsException details (message and stack trace) when the password is invalid. | Demonstrate catching both Aspose.Cells-specific CellsException and generic Exception while loading a protected Excel file, and output comprehensive error information to the console.
// Common Searches: aspocells c# load encrypted xlsx file with wrong password and get detailed error | how to handle CellsException when opening password protected Excel workbook in .NET | log stack trace for failed workbook load using Aspose.Cells LoadOptions password
// Tags: load encrypted workbook using Aspose.Cells LoadOptions | catch CellsException authentication error | log workbook load failure details | verify excel file existence before opening | Aspose.Cells password handling in .NET

using System;
using System.IO;
using Aspose.Cells;

// // Checks for the presence of "protected.xlsx", attempts to open it with a supplied password via Aspose.Cells LoadOptions, and captures both CellsException and generic Exception, printing the error message and stack trace.
class Program
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string filePath = "protected.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // The password to try (replace with the correct one as needed)
        string password = "incorrectPassword";

        try
        {
            // Configure load options with the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Attempt to load the workbook
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., incorrect password)
            Console.WriteLine("Aspose.Cells error occurred while loading the workbook.");
            Console.WriteLine($"Message      : {ex.Message}");
            Console.WriteLine($"Stack Trace  : {ex.StackTrace}");
        }
        catch (Exception ex)
        {
            // General exception handling for any other errors
            Console.WriteLine("An unexpected error occurred while loading the workbook.");
            Console.WriteLine($"Message      : {ex.Message}");
            Console.WriteLine($"Stack Trace  : {ex.StackTrace}");
        }
    }
}
