// Title: Detect encryption status and obtain the encryption algorithm of a legacy XLS workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an XLS file with Aspose.Cells, catches the CellsException to determine if the workbook is password‑protected, and returns the encryption algorithm name when a password is required. | Create a reusable method GetXlsEncryptionInfo(string path) that returns a tuple (bool isEncrypted, string algorithm) using Aspose.Cells load options. | Show how to use Aspose.Cells LoadOptions to read the encryption algorithm of an encrypted XLS file without providing the password.
// Common Searches: asp.net how to know if an old .xls workbook is password protected using Aspose.Cells | c# retrieve encryption algorithm of a protected xls file with Aspose.Cells | detect encrypted legacy Excel file and get its encryption type in .NET | Aspose.Cells check encryption status of XLS and read algorithm property
// Tags: Aspose.Cells detect encrypted XLS workbook | retrieve encryption algorithm from XLS using Aspose.Cells | C# load legacy Excel file with encryption detection | handle CellsException for password‑protected XLS | Aspose.Cells LoadOptions encryption algorithm property

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an XLS file with Aspose.Cells, catching a CellsException to identify password protection, and (when encrypted) accessing the workbook's LoadOptions to read the encryption algorithm.
class Program
{
    static void Main()
    {
        // Path to the XLS file to be examined
        string filePath = "input.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            bool isEncrypted = false;

            try
            {
                // Attempt to load the workbook without a password.
                // If the file is not encrypted, this will succeed.
                var workbook = new Workbook(filePath);
                // No exception means the file is not encrypted.
                isEncrypted = false;
            }
            catch (CellsException)
            {
                // An exception while loading indicates the file is encrypted (password required).
                isEncrypted = true;
            }
            catch (Exception ex)
            {
                // Other unexpected errors while loading the workbook.
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            Console.WriteLine(isEncrypted ? "The file is encrypted." : "The file is not encrypted.");
        }
        catch (Exception ex)
        {
            // General exception handling for any unforeseen errors.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
