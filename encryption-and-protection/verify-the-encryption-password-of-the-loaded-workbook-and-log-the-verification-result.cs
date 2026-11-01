// Title: Verify the password of an encrypted Excel workbook with Aspose.Cells LoadOptions in C#
// Common Searches: C# Aspose.Cells check if an encrypted Excel file password is correct | how to catch CellsException for wrong password when loading a protected workbook | validate encrypted XLSX password using LoadOptions.Password in .NET | determine whether an Excel workbook is password protected with Aspose.Cells | log the result of Excel password verification in a C# console app
// Tags: Aspose.Cells password verification | LoadOptions.Password usage | CellsException invalid password handling | encrypted XLSX loading C# | console log verification result

using System;
using System.IO;
using Aspose.Cells;

// The example loads an encrypted XLSX file using Aspose.Cells LoadOptions with a supplied password. If the workbook loads without throwing a CellsException indicating an invalid password, the password is confirmed as correct; otherwise it is flagged as incorrect. The verification outcome is written to the console.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Ensure the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Password to test
        string passwordToTest = "MySecretPassword";

        bool isPasswordCorrect = false;

        try
        {
            // Load the workbook with the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = passwordToTest
            };

            // If loading succeeds, the password is correct
            Workbook workbook = new Workbook(filePath, loadOptions);
            isPasswordCorrect = true;
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for invalid password.
            // Older versions may not expose ErrorCode, so inspect the message.
            if (ex.Message != null && ex.Message.IndexOf("Invalid password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isPasswordCorrect = false;
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Handle other possible errors (e.g., corrupted file)
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Log the verification result
        if (isPasswordCorrect)
        {
            Console.WriteLine("Password verification succeeded: the password is correct.");
        }
        else
        {
            Console.WriteLine("Password verification failed: the password is incorrect.");
        }
    }
}
