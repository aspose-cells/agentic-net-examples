// Title: Load a password‑protected Excel workbook with Aspose.Cells for .NET and handle incorrect‑password exceptions
// Description: Demonstrates how to open a password‑protected Excel file using Aspose.Cells LoadOptions, catch the CellsException with ExceptionType.IncorrectPassword, and log the exception code and message while also handling any other loading errors.
// Keywords: Aspose.Cells | .NET | password protected Excel | LoadOptions | CellsException | IncorrectPassword | authentication error logging | workbook loading error | exception code | exception message
// Common Searches: Aspose.Cells open password protected workbook .NET | catch IncorrectPassword exception Aspose.Cells | log authentication failure Aspose.Cells | exception code for wrong Excel password | how to handle wrong password when loading Excel with Aspose.Cells
// Developer Intent: Open a password‑protected Excel file with Aspose.Cells, detect an invalid password, and record detailed error information.
// Use Cases: Validate a user‑entered password before loading a workbook and show a friendly error if it fails. | Send exception code and message to a monitoring or logging service when authentication is rejected. | Trigger a fallback routine (e.g., request a new password or load an unprotected template) after catching an IncorrectPassword error.
// AI Prompts: Generate a C# snippet that loads a password‑protected Excel workbook with Aspose.Cells, catches CellsException for IncorrectPassword, and writes the exception details to a log file. | Show how to differentiate between IncorrectPassword and other load errors in Aspose.Cells and implement a retry loop for password entry.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordLoadDemo
{
    // Demonstrates how to open a password‑protected Excel file using Aspose.Cells LoadOptions, catch the CellsException with ExceptionType.IncorrectPassword, and log the exception code and message while also handling any other loading errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the password‑protected Excel file
            string filePath = "protected.xlsx";

            // Password to attempt opening the file (replace with the correct one as needed)
            string password = "wrongPassword";

            // Configure load options with the supplied password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = password;

            try
            {
                // Attempt to load the workbook using the load options
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully.");

                // Example verification: read a cell value
                Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);
            }
            // Catch Aspose.Cells specific exception for incorrect password
            catch (CellsException ex) when (ex.Code == ExceptionType.IncorrectPassword)
            {
                Console.WriteLine("Authentication failed: Incorrect password.");
                Console.WriteLine($"Exception Code: {ex.Code}");
                Console.WriteLine($"Exception Message: {ex.Message}");
            }
            // Catch any other unexpected exceptions
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred while loading the workbook.");
                Console.WriteLine($"Exception Message: {ex.Message}");
            }
        }
    }
}
