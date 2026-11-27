// Title: Determine if an Excel workbook is encrypted and flag unsupported encryption algorithms using Aspose.Cells for .NET
// Common Searches: how to check if an Excel file is password protected with Aspose.Cells in C# | Aspose.Cells detect outdated encryption algorithm in an .xlsx workbook | C# code to load encrypted Excel file and identify unsupported encryption | flag Excel workbook for migration when using deprecated encryption with Aspose.Cells | load encrypted workbook with fallback password using Aspose.Cells C#
// Tags: detect workbook encryption Aspose.Cells | check Excel encryption status C# | flag deprecated encryption algorithm .NET | load encrypted xlsx with password Aspose.Cells | Workbook.Settings.IsEncrypted usage

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, supplies a password when required, uses Workbook.Settings.IsEncrypted to determine if the workbook is encrypted, and notes that the specific encryption algorithm cannot be identified with the current Aspose.Cells version, allowing developers to flag files that may need migration to supported encryption methods.
class EncryptionChecker
{
    static void Main()
    {
        // Path to the workbook to be inspected
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // LoadOptions without a password initially
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (CellsException)
        {
            // Assume the exception is due to missing/incorrect password
            // Provide the password here (replace with the actual password or retrieve it securely)
            loadOptions.Password = "YourPasswordHere";

            try
            {
                // Reload the workbook with the supplied password
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open workbook with password: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Determine if the workbook is encrypted
        bool isEncrypted = false;
        try
        {
            isEncrypted = workbook.Settings.IsEncrypted;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unable to determine encryption status: {ex.Message}");
        }

        if (!isEncrypted)
        {
            Console.WriteLine("Workbook is not encrypted.");
            return;
        }

        // Aspose.Cells does not expose the encryption algorithm directly in all versions.
        // If needed, additional logic can be added here using newer APIs when available.
        Console.WriteLine("Workbook is encrypted. Encryption algorithm detection is not supported in this version of Aspose.Cells.");
    }
}
