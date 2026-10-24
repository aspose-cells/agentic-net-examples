// Title: Check if an Excel workbook is encrypted and log a recommendation to use stronger AES256 encryption with Aspose.Cells for .NET
// Common Searches: how to determine if an Excel file is password protected using Aspose.Cells C# | Aspose.Cells C# detect encrypted workbook and suggest stronger encryption | C# read workbook encryption status with Aspose.Cells LoadOptions | log recommendation for upgrading Excel encryption to AES256 in .NET | using workbook.Settings.IsEncrypted property to check Excel protection
// Tags: Aspose.Cells encrypted workbook detection | LoadOptions usage for encrypted Excel files | C# recommendation for AES256 Excel encryption | Workbook.Settings.IsEncrypted property example | log encryption upgrade advice .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells LoadOptions, checks the Settings.IsEncrypted flag, and writes a console message recommending a stronger encryption algorithm such as AES256 when the file is encrypted. It also includes file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        // Path to the workbook to be inspected
        string workbookPath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Prepare load options – this allows us to read encryption information
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // If the workbook is password‑protected, set the password here
            // loadOptions.Password = "yourPassword";

            // Load the workbook using the load options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // Determine whether the workbook is encrypted
            bool isEncrypted = workbook.Settings.IsEncrypted;

            if (isEncrypted)
            {
                // Note: Aspose.Cells does not expose the specific encryption algorithm.
                // Recommend using a strong algorithm when creating/encrypting workbooks.
                Console.WriteLine("The workbook is encrypted. Consider using a strong encryption algorithm (e.g., AES256) when protecting workbooks.");
            }
            else
            {
                Console.WriteLine("The workbook is not encrypted.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
