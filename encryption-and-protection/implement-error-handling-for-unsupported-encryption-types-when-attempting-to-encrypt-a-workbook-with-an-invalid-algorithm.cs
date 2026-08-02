// Title: Handle Unsupported Encryption Algorithm Exception in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to catch a CellsException with the UnsupportedFeature code when an invalid EncryptionType is passed to Workbook.SetEncryptionOptions. The example creates a workbook, adds data, applies a password, forces an unsupported algorithm, and provides specific and generic error handling.
// Keywords: Aspose.Cells | C# | Workbook encryption | EncryptionType invalid | SetEncryptionOptions | UnsupportedFeature | CellsException handling | error handling encryption | invalid algorithm | exception handling
// Common Searches: Aspose.Cells catch unsupported encryption type | SetEncryptionOptions invalid algorithm exception | CellsException UnsupportedFeature example | C# workbook encryption error handling | how to validate EncryptionType Aspose.Cells
// Developer Intent: Implement robust exception handling to detect and respond when an unsupported encryption algorithm is supplied to Aspose.Cells workbook encryption.
// Use Cases: Show a user‑friendly message when an invalid EncryptionType is used. | Log CellsException details (code and message) for troubleshooting encryption failures. | Automatically switch to a supported encryption algorithm after catching the UnsupportedFeature exception.
// AI Prompts: Generate C# code that validates EncryptionType before calling SetEncryptionOptions in Aspose.Cells. | Provide a fallback routine that selects a default supported encryption algorithm after catching an UnsupportedFeature exception. | Explain the different CellsException codes related to workbook encryption in Aspose.Cells and how to handle each.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates how to catch a CellsException with the UnsupportedFeature code when an invalid EncryptionType is passed to Workbook.SetEncryptionOptions. The example creates a workbook, adds data, applies a password, forces an unsupported algorithm, and provides specific and generic error handling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encryption Test");

            // Set a password for the workbook
            workbook.Settings.Password = "myPassword";

            // Attempt to set encryption options with an unsupported algorithm
            // Here we cast an undefined integer to EncryptionType to simulate an invalid value
            EncryptionType invalidEncryption = (EncryptionType)99;
            int keyLength = 128; // Valid key length, but encryption type is invalid

            try
            {
                // This call should raise an exception because the encryption type is not supported
                workbook.SetEncryptionOptions(invalidEncryption, keyLength);

                // If no exception, save the workbook (lifecycle save)
                workbook.Save("EncryptedWorkbook.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
            {
                // Handle the specific case where the encryption algorithm is unsupported
                Console.WriteLine("Error: The selected encryption type is not supported.");
                Console.WriteLine($"Exception Code: {ex.Code}");
                Console.WriteLine($"Message: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General fallback for any other unexpected errors
                Console.WriteLine("An unexpected error occurred while encrypting the workbook.");
                Console.WriteLine($"Message: {ex.Message}");
            }
        }
    }
}
