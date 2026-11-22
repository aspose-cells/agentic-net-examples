// Title: Detect if an .xlsx workbook is encrypted using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells determine whether an Excel file is password protected without providing a password | How to test if a .xlsx file is encrypted before loading with Aspose.Cells | Programmatically check encryption status of Office Open XML workbook using Aspose.Cells | Aspose.Cells load options to identify encrypted Excel workbook in .NET | Identify encrypted Excel files and handle exceptions with Aspose.Cells C#
// Tags: Aspose.Cells workbook encryption detection | C# detect encrypted xlsx file | Aspose.Cells loadoptions password handling | Office Open XML encryption status check | Excel file encryption detection .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a specified .xlsx file with Aspose.Cells, attempts to open it without a password, and treats any load exception as an indication that the workbook is encrypted, then reports the encryption status and shows how to use LoadOptions with a password for further processing.
class Program
{
    static void Main()
    {
        // Path to the Office Open XML file (e.g., .xlsx) that may be encrypted
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            bool isEncrypted = false;

            // Attempt to load the workbook without a password.
            // If an exception occurs, assume the file is encrypted.
            try
            {
                using (var wbCheck = new Workbook(filePath))
                {
                    // Loaded successfully, not encrypted.
                }
            }
            catch
            {
                isEncrypted = true;
            }

            if (isEncrypted)
            {
                Console.WriteLine("The file is encrypted.");
                // If you need to work with the encrypted file, provide the password in LoadOptions
                // Example (uncomment and set the password):
                // LoadOptions lo = new LoadOptions { Password = "yourPassword" };
                // Workbook wbEncrypted = new Workbook(filePath, lo);
            }
            else
            {
                // Load the workbook normally (no password required)
                LoadOptions loadOptions = new LoadOptions { Password = null };
                Workbook wb = new Workbook(filePath, loadOptions);
                Console.WriteLine("The file is NOT encrypted.");
                // Perform further processing with 'wb' as needed
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
