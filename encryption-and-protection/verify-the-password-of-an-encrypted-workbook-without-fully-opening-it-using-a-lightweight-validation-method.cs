// Title: Validate an encrypted Excel workbook password with Aspose.Cells in C# without fully loading the file
// Common Searches: C# Aspose.Cells check Excel file password without opening workbook | how to test password of encrypted .xlsx using LoadOptions Aspose.Cells | lightweight password verification for protected Excel workbook in .NET | validate workbook password without reading all cells Aspose.Cells | detect if Excel file is password protected using Aspose.Cells C#
// Tags: Aspose.Cells workbook password validation | LoadOptions password check encrypted XLSX | C# lightweight Excel password verification | CellsException handling invalid workbook password | verify Excel file protection without full load

using Aspose.Cells;
using System;
using System.IO;

// The example shows how to verify the password of an encrypted .xlsx workbook by loading it with LoadOptions that include the supplied password. If the password is correct the Workbook loads and true is returned; otherwise a CellsException is caught and false is returned, enabling a lightweight password check without processing the entire workbook.
class PasswordValidator
{
    static void Main()
    {
        try
        {
            // Path to the encrypted workbook file
            string workbookPath = @"C:\Path\To\EncryptedWorkbook.xlsx";

            // Password to validate
            string passwordToTest = "MySecretPassword";

            // Ensure the file exists before attempting validation
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Validate the password by attempting to load the workbook with the supplied password
            bool isPasswordValid = ValidatePassword(workbookPath, passwordToTest);

            // Output the result
            Console.WriteLine(isPasswordValid
                ? "Password is correct."
                : "Password is incorrect or the file is not password protected.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static bool ValidatePassword(string filePath, string password)
    {
        try
        {
            // Set load options with the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Attempt to load the workbook; if the password is wrong, an exception will be thrown
            using (Workbook wb = new Workbook(filePath, loadOptions))
            {
                // If we reach this point, the password is correct
                return true;
            }
        }
        catch (CellsException)
        {
            // Incorrect password or file not password protected
            return false;
        }
        catch (Exception)
        {
            // Any other error (e.g., corrupted file) is treated as validation failure
            return false;
        }
    }
}
