// Title: How to open an encrypted Excel workbook with an incorrect password and retrieve Aspose.Cells exception information in C#
// AI Prompts: Load a password‑protected .xlsx file using Aspose.Cells with an invalid password and print the CellsException message and stack trace. | Demonstrate catching Aspose.Cells CellsException when opening an encrypted workbook with a wrong password in a C# console application. | Show how to configure LoadOptions with an incorrect password and handle the resulting error from Aspose.Cells.
// Common Searches: Aspose.Cells C# how to handle wrong password error when opening encrypted Excel file | retrieve exception details for invalid workbook password using Aspose.Cells LoadOptions | C# example catching CellsException for password‑protected .xlsx file
// Tags: catch CellsException for invalid workbook password | load encrypted .xlsx with Aspose.Cells LoadOptions | handle wrong password error in Aspose.Cells | exception handling for password‑protected Excel in C# | Aspose.Cells workbook decryption failure handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks for the presence of an encrypted Excel file, attempts to open it with Aspose.Cells using LoadOptions that contain an incorrect password, and captures both CellsException and generic Exception to display detailed error messages and stack traces.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Incorrect password
            string wrongPassword = "WrongPassword";

            // Load options with the wrong password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = wrongPassword
            };

            try
            {
                // Attempt to open the workbook
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook opened successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Aspose.Cells specific exception handling
                Console.WriteLine("Failed to open workbook:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
    }
}
