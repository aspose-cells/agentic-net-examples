// Title: Load a password‑protected Excel workbook with Aspose.Cells for .NET and log authentication errors
// Common Searches: Aspose.Cells C# load encrypted XLSX file using password and catch incorrect password exception | catch CellsException when opening a protected Excel workbook with Aspose.Cells | log detailed error information for failed workbook load in Aspose.Cells .NET | verify Excel file exists before loading with Aspose.Cells and a password | use LoadOptions.Password to open a password‑protected workbook in Aspose.Cells
// Tags: load password protected xlsx with Aspose.Cells | CellsException handling for workbook authentication | Aspose.Cells LoadOptions password configuration | detailed error logging for workbook load failures | file existence check before Aspose.Cells workbook load

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to open a password‑protected Excel file using Aspose.Cells for .NET. It checks that the file exists, configures LoadOptions with the supplied password, and attempts to create a Workbook. Specific CellsException handling captures authentication errors such as an incorrect password, while a generic catch logs any other unexpected exceptions. The LogError method records a timestamp, exception type, message, stack trace, and inner‑exception details for comprehensive troubleshooting.
class Program
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string filePath = @"C:\Data\ProtectedWorkbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        // The password to open the workbook (replace with the correct one or test with a wrong password)
        string password = "WrongPassword";

        // Configure load options with the supplied password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };

        try
        {
            // Attempt to load the workbook using the password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, you can work with the workbook here
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells specific errors (e.g., incorrect password)
            LogError("Aspose.Cells error occurred while loading the workbook.", ex);
        }
        catch (Exception ex)
        {
            // Any other unexpected errors
            LogError("An unexpected error occurred.", ex);
        }
    }

    /// <param name="message">Custom message describing the context.</param>
    /// <param name="ex">The exception to log.</param>
    static void LogError(string message, Exception ex)
    {
        Console.WriteLine("=== Error Log Start ===");
        Console.WriteLine($"Timestamp   : {DateTime.UtcNow:O}");
        Console.WriteLine($"Message     : {message}");
        Console.WriteLine($"Exception   : {ex.GetType().FullName}");
        Console.WriteLine($"Message     : {ex.Message}");
        Console.WriteLine($"Stack Trace : {ex.StackTrace}");

        if (ex is CellsException ce)
        {
            // CellsException may contain additional information; output what is available
            Console.WriteLine($"Error Code  : {ce.Message}");
        }

        if (ex.InnerException != null)
        {
            Console.WriteLine("--- Inner Exception ---");
            Console.WriteLine($"Type        : {ex.InnerException.GetType().FullName}");
            Console.WriteLine($"Message     : {ex.InnerException.Message}");
            Console.WriteLine($"Stack Trace : {ex.InnerException.StackTrace}");
        }

        Console.WriteLine("=== Error Log End ===");
    }
}
