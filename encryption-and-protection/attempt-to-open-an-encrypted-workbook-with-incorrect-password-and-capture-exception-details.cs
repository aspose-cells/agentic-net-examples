// Title: Catch and display exception details when loading an encrypted Excel workbook with an incorrect password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to handle wrong password exception when opening an encrypted workbook | catch exception for invalid password in LoadOptions with Aspose.Cells | retrieve exception message and stack trace when loading a password‑protected XLSX using Aspose.Cells | C# load encrypted Excel file with wrong password and log error details Aspose.Cells | Aspose.Cells workbook load fails with invalid password how to capture exception
// Tags: load encrypted xlsx with incorrect password Aspose.Cells | exception handling for workbook decryption failure | Aspose.Cells LoadOptions password error handling | capture stack trace for failed workbook load | invalid password handling in Aspose.Cells .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example sets an incorrect password in LoadOptions, attempts to open an encrypted XLSX workbook with Aspose.Cells, catches any thrown exception, and prints the failure message, exception type, message, and stack trace.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the encrypted workbook file
            string filePath = "encrypted.xlsx";

            // Set an incorrect password in LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "incorrectPassword"
            };

            try
            {
                // Attempt to load the workbook with the wrong password
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Capture and display exception details
                Console.WriteLine("Failed to open the encrypted workbook.");
                Console.WriteLine("Exception Message: " + ex.Message);
                Console.WriteLine("Exception Type: " + ex.GetType().FullName);
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
