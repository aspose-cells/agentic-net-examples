// Title: Handle Unsupported Encryption Types When Encrypting a Workbook with Aspose.Cells for .NET
// Description: Shows how to catch a CellsException (UnsupportedFeature) caused by an invalid EncryptionType passed to SetEncryptionOptions, log the error, and save the workbook with default protection.
// Keywords: Aspose.Cells encryption error handling | Unsupported EncryptionType exception | SetEncryptionOptions invalid algorithm | C# workbook password protection | CellsException UnsupportedFeature
// Common Searches: Aspose.Cells catch unsupported encryption algorithm | Exception thrown for invalid EncryptionType in .NET | Validate EncryptionType before SetEncryptionOptions | How to fallback to default encryption in Aspose.Cells | C# encrypt workbook with Aspose.Cells error handling
// Developer Intent: Implement robust exception handling to manage unsupported encryption algorithms when applying workbook protection.
// Use Cases: Check the EncryptionType value against the enum before calling SetEncryptionOptions to prevent runtime errors. | Log detailed information from a CellsException (UnsupportedFeature) and revert to default encryption settings. | Encapsulate encryption logic in a reusable method that returns success status and handles both specific and generic exceptions.
// AI Prompts: Create a helper method that verifies an integer maps to a valid EncryptionType and applies encryption with proper error handling. | Show how to log CellsException details for UnsupportedFeature while providing a fallback encryption configuration in Aspose.Cells. | Write code that lists all supported EncryptionType values, lets a user select one, and gracefully handles invalid selections.

using System;
using Aspose.Cells;

// Shows how to catch a CellsException (UnsupportedFeature) caused by an invalid EncryptionType passed to SetEncryptionOptions, log the error, and save the workbook with default protection.
class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

        // Set a password for encryption
        workbook.Settings.Password = "myPassword";

        // Attempt to apply an unsupported encryption algorithm
        try
        {
            // Cast an invalid integer to EncryptionType to simulate an unsupported type
            EncryptionType unsupportedType = (EncryptionType)999;
            workbook.SetEncryptionOptions(unsupportedType, 128);
            Console.WriteLine("Encryption options applied successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
        {
            // Specific handling for unsupported encryption features
            Console.WriteLine("Unsupported encryption type: " + ex.Message);
        }
        catch (Exception ex)
        {
            // General fallback for any other errors
            Console.WriteLine("Error applying encryption options: " + ex.Message);
        }

        // Save the workbook; if SetEncryptionOptions failed, the workbook will be saved with default encryption
        workbook.Save("EncryptedWorkbook.xlsx");
    }
}
