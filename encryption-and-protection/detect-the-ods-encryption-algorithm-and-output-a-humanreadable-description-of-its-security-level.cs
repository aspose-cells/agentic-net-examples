// Title: Detect ODS workbook encryption status and obtain its security description with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells example to verify if an ODS file is password protected | how to programmatically determine encryption status of an ODS spreadsheet using Aspose.Cells | catch CellsException to identify encrypted ODS workbook in .NET | retrieve security level description of an ODS file with Aspose.Cells API | load ODS file with LoadOptions and check for encryption in C#
// Tags: Aspose.Cells ODS encryption detection | C# load ODS workbook with LoadOptions | identify password protected ODS file Aspose.Cells | handle CellsException for encrypted ODS workbook | report ODS security level programmatically

using System;
using System.IO;
using Aspose.Cells;

// The sample checks whether an ODS workbook is encrypted by attempting to load it with Aspose.Cells LoadOptions. If loading succeeds, it reports no encryption; if a CellsException is thrown, it flags the file as encrypted and notes that the specific algorithm isn’t exposed, then outputs a human‑readable security description.
class OdsEncryptionDetector
{
    static void Main()
    {
        // Path to the ODS file to analyze
        string odsPath = "sample.ods";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(odsPath))
        {
            Console.WriteLine($"Error: The file \"{odsPath}\" was not found.");
            return;
        }

        // Load options for ODS format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods);

        bool isEncrypted = false;
        string description = string.Empty;

        try
        {
            // Attempt to load the workbook; if it succeeds, the file is not encrypted
            Workbook workbook = new Workbook(odsPath, loadOptions);
            description = "No encryption applied.";
        }
        catch (CellsException cex)
        {
            // Aspose.Cells throws CellsException for password‑protected files
            // Since specific error codes are not exposed in newer versions,
            // we treat any CellsException during load as an indication of encryption.
            isEncrypted = true;
            description = "Encrypted ODS file (encryption algorithm not exposed by API).";
        }
        catch (Exception ex)
        {
            // Handle other unexpected errors (e.g., corrupted file, unsupported format)
            Console.WriteLine($"An error occurred while processing the file: {ex.Message}");
            return;
        }

        // Output the result
        Console.WriteLine($"Encryption detected: {isEncrypted}");
        Console.WriteLine($"Security level: {description}");
    }
}
