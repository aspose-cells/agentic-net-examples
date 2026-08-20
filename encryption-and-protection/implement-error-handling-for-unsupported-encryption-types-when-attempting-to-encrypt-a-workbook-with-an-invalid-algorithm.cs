// Title: C# – Handle Unsupported Encryption Types When Encrypting an Aspose.Cells Workbook
// Description: Shows how to catch the CellsException (UnsupportedFeature) that Aspose.Cells throws if SetEncryptionOptions receives an invalid EncryptionType, and how to log the error or switch to a supported algorithm.
// Keywords: Aspose.Cells C# encryption error handling | unsupported EncryptionType | CellsException UnsupportedFeature | SetEncryptionOptions exception | workbook password protection .NET | Excel file encryption fallback | Aspose.Cells security API | invalid encryption algorithm handling
// Common Searches: Aspose.Cells catch unsupported encryption type | SetEncryptionOptions throws CellsException | how to validate EncryptionType before encrypting workbook | C# encrypt Excel file with Aspose.Cells error handling | supported encryption algorithms Aspose.Cells .NET
// Developer Intent: Add reliable error handling for invalid encryption algorithms when applying password protection to a workbook.
// Use Cases: Check the requested EncryptionType against Aspose.Cells' enum values and throw a custom exception if it is not supported. | Log detailed information when an UnsupportedFeature exception occurs and inform the user about the unavailable algorithm. | Automatically fall back to a default supported encryption (e.g., AES128) when the supplied type is invalid.
// AI Prompts: Generate C# code that validates an EncryptionType before calling SetEncryptionOptions and raises a custom InvalidEncryptionException for unsupported values. | Create a helper method that maps supported EncryptionType values to key sizes, logs an error for unknown enums, and returns a default AES128 option. | Write a try‑catch example that captures CellsException with code UnsupportedFeature, logs the message and stack trace, and then retries encryption with a fallback algorithm.

using System;
using Aspose.Cells;

// Shows how to catch the CellsException (UnsupportedFeature) that Aspose.Cells throws if SetEncryptionOptions receives an invalid EncryptionType, and how to log the error or switch to a supported algorithm.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

        // Set a password for encryption
        workbook.Settings.Password = "MySecretPassword";

        try
        {
            // Simulate an unsupported encryption algorithm by casting an undefined enum value
            EncryptionType unsupportedEncryption = (EncryptionType)99;
            workbook.SetEncryptionOptions(unsupportedEncryption, 128);

            // Save the workbook (this line will not be reached if the encryption type is unsupported)
            workbook.Save("EncryptedWorkbook.xlsx");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
        {
            // Handle the specific case where the encryption type is not supported
            Console.WriteLine("Error: The specified encryption type is not supported.");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
