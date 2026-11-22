// Title: Check an encrypted Excel workbook’s password with Aspose.Cells for .NET (C#)
// Common Searches: how to validate an Excel .xlsx password using Aspose.Cells LoadOptions in C# | Aspose.Cells C# example for testing workbook encryption password | catch CellsException when opening password‑protected Excel file with Aspose.Cells | verify Excel file password before loading with Aspose.Cells .NET API | determine if supplied password opens encrypted workbook using Aspose.Cells
// Tags: Aspose.Cells LoadOptions password verification | C# password‑protected Excel workbook loading | CellsException handling for invalid workbook password | encrypted .xlsx file opening with Aspose.Cells | Excel encryption password check .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordCheck
{
    // The sample loads an encrypted Excel file using LoadOptions with a supplied password; successful loading confirms the password, while a CellsException indicates an incorrect password and other exceptions are reported as errors.
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Password to verify
            string passwordToTest = "MySecretPassword";

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // LoadOptions with the password to attempt opening the workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = passwordToTest
                };

                // Attempt to load the workbook using the supplied password
                Workbook workbook = new Workbook(filePath, loadOptions);

                // If no exception is thrown, the password is correct
                Console.WriteLine("Password verification succeeded: the password is correct.");
            }
            catch (CellsException)
            {
                // Any CellsException during load typically indicates an incorrect password
                Console.WriteLine("Password verification failed: the password is incorrect or the file is invalid.");
            }
            catch (Exception ex)
            {
                // Any other errors (e.g., file corrupted, access issues)
                Console.WriteLine($"An error occurred while verifying the password: {ex.Message}");
            }
        }
    }
}
