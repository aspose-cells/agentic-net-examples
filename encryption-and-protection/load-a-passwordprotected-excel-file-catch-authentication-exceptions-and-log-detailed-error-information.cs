// Title: Open a password‑protected XLSX workbook with Aspose.Cells for .NET and log authentication failures
// Common Searches: asp.net load encrypted Excel file using Aspose.Cells and handle wrong password exception | c# catch CellsException when opening password protected workbook with Aspose.Cells | log detailed error information for failed Excel decryption in Aspose.Cells | verify file existence before loading password protected XLSX with Aspose.Cells | how to use LoadOptions.Password in Aspose.Cells to open a protected workbook
// Tags: Aspose.Cells LoadOptions password protection | C# CellsException handling for encrypted Excel | Aspose.Cells workbook loading error logging | password protected XLSX opening with Aspose.Cells | exception detail logging in .NET Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that the XLSX file exists, creates a LoadOptions object with the workbook password, and loads the file using Aspose.Cells. It catches CellsException to detect an incorrect password, handles any other exceptions, and writes detailed error information (message, source, stack trace, inner exceptions) to the console via a helper method.
class Program
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string filePath = "protected.xlsx";

        // Password to open the file
        string password = "myPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.Error.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Configure load options with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
            // Additional processing can be done here
        }
        // Handle incorrect password scenario
        catch (CellsException ex)
        {
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.Error.WriteLine("Authentication failed: Incorrect password.");
            }
            else
            {
                Console.Error.WriteLine("A CellsException occurred while loading the workbook.");
            }
            LogExceptionDetails(ex);
        }
        // Catch any other exceptions that may occur during loading
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred while loading the workbook.");
            LogExceptionDetails(ex);
        }
    }

    // Helper method to log detailed error information
    static void LogExceptionDetails(Exception ex)
    {
        Console.Error.WriteLine($"Message   : {ex.Message}");
        Console.Error.WriteLine($"Source    : {ex.Source}");
        Console.Error.WriteLine($"StackTrace: {ex.StackTrace}");
        if (ex.InnerException != null)
        {
            Console.Error.WriteLine("Inner Exception:");
            LogExceptionDetails(ex.InnerException);
        }
    }
}
