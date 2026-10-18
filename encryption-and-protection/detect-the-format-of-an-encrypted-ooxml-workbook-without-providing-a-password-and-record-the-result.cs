// Title: Detect encrypted OOXML Excel workbook format without a password using Aspose.Cells for .NET (C#)
// Common Searches: how to check if an .xlsx file is password protected with Aspose.Cells in C# | detect workbook file format and encryption status using LoadOptions Auto in Aspose.Cells | C# code to determine if an Excel file requires a password without opening it | Aspose.Cells load encrypted workbook without password exception handling example | identify OOXML workbook format and encryption flag in a .NET application
// Tags: encrypted workbook detection Aspose.Cells | LoadOptions Auto format detection C# | CellsException password handling Aspose.Cells | Workbook.FileFormat extraction C# | load Excel file without password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file path, creates LoadOptions with LoadFormat.Auto and a null password, attempts to open the workbook using Aspose.Cells, catches a CellsException containing the word "Password" to flag encryption, and then outputs the detected file format (from Workbook.FileFormat or LoadOptions.LoadFormat) together with the encryption status.
class DetectEncryptedWorkbook
{
    static void Main()
    {
        // Path to the workbook (could be .xlsx, .xlsm, etc.).
        string workbookPath = "encryptedWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException.
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        // LoadOptions with Auto format detection and no password.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
        {
            Password = null
        };

        bool isEncrypted = false;
        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password.
            workbook = new Workbook(workbookPath, loadOptions);
        }
        catch (CellsException cex)
        {
            // If the exception message indicates a password is required, treat it as encrypted.
            if (!string.IsNullOrEmpty(cex.Message) && cex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isEncrypted = true;
            }
            else
            {
                // Other CellsException errors (corrupt file, unsupported format, etc.).
                Console.WriteLine($"Error loading workbook: {cex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Non‑CellsException errors.
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Determine the detected format.
        string detectedFormat = workbook != null
            ? workbook.FileFormat.ToString()
            : loadOptions.LoadFormat.ToString();

        // Output the results.
        Console.WriteLine("Detected Format: " + detectedFormat);
        Console.WriteLine("Is Encrypted: " + isEncrypted);
    }
}
