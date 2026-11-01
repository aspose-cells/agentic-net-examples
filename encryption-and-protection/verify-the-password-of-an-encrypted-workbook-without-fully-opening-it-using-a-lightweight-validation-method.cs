// Title: Check an Excel workbook password with Aspose.Cells LoadOptions without fully loading the file (C#)
// Common Searches: how to test Excel file password with Aspose.Cells without opening the workbook in C# | C# verify encrypted .xlsx password using Aspose.Cells LoadOptions | lightweight password validation for protected Excel workbook Aspose.Cells | Aspose.Cells check workbook password without reading all data | programmatically validate workbook encryption password with Aspose.Cells .NET
// Tags: Aspose.Cells LoadOptions password validation | encrypted .xlsx password check C# | validate Excel workbook protection Aspose.Cells | lightweight workbook opening Aspose.Cells | C# password verification for protected Excel file

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to use Aspose.Cells LoadOptions with a supplied password to attempt opening an encrypted workbook. If the Workbook loads successfully, the password is correct; otherwise, an exception indicates an invalid password, enabling password verification without fully processing the file.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string workbookPath = @"C:\Path\To\EncryptedWorkbook.xlsx";

        // Password to validate
        string passwordToTest = "MySecretPassword";

        // Ensure the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine("Error: Workbook file not found at the specified path.");
            return;
        }

        bool isPasswordValid = false;

        try
        {
            // Set load options with the password to attempt opening the workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = passwordToTest
            };

            // Attempt to load the workbook; if the password is incorrect, an exception will be thrown
            using (Workbook wb = new Workbook(workbookPath, loadOptions))
            {
                // If loading succeeds, the password is valid
                isPasswordValid = true;
            }
        }
        catch (Exception)
        {
            // Loading failed – most likely due to an invalid password
            isPasswordValid = false;
        }

        // Output the result
        Console.WriteLine(isPasswordValid
            ? "Password is correct."
            : "Password is invalid.");
    }
}
