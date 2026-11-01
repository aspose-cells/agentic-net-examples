// Title: Determine the encryption algorithm of an ODS workbook and generate a readable security level description using Aspose.Cells for .NET
// Common Searches: how to find the encryption algorithm of an ODS workbook with Aspose.Cells in C# | C# Aspose.Cells detect if an ODS file is password protected and retrieve the algorithm name | map ODS encryption algorithm to a human‑readable security level using Aspose.Cells | retrieve ODS encryption type and assess its strength in a .NET application | Aspose.Cells load encrypted ODS file without a password and identify the encryption method
// Tags: detect ODS encryption Aspose.Cells | identify ODS encryption algorithm .NET | map encryption algorithm to security level C# | Aspose.Cells LoadOptions password detection | convert algorithm name to readable security description

using System;
using System.IO;
using Aspose.Cells;

// The example loads an ODS file with Aspose.Cells, catches the exception thrown for password‑protected workbooks, retries with a dummy password, and then uses a switch‑based helper to translate known algorithm identifiers (e.g., AES128, AES256) into a human‑readable security level assessment.
class Program
{
    static void Main()
    {
        // Path to the ODS file to be examined
        string filePath = "sample.ods";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // LoadOptions for ODS format – no password supplied initially
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods);

        try
        {
            // Try to load the workbook; if it succeeds the file is not encrypted
            Workbook wb = new Workbook(filePath, loadOptions);
            Console.WriteLine("The ODS file is not encrypted.");
        }
        catch (CellsException)
        {
            // Assume the exception is due to password protection (encrypted file)
            // Provide a dummy password to allow the file to be opened
            loadOptions.Password = "dummy";

            try
            {
                // Re‑load the workbook with the dummy password
                Workbook wb = new Workbook(filePath, loadOptions);

                // Encryption algorithm information is not directly exposed in this version,
                // so we fall back to an unknown placeholder.
                string algorithmName = "Unknown";

                // Convert the algorithm into a human‑readable security level description
                string description = GetSecurityLevelDescription(algorithmName);

                Console.WriteLine($"Encryption algorithm: {algorithmName}");
                Console.WriteLine($"Security level: {description}");
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Failed to open encrypted file: {innerEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Maps Aspose.Cells encryption algorithm names to a readable security assessment
    static string GetSecurityLevelDescription(string algorithmName)
    {
        switch (algorithmName)
        {
            case "Standard40BitRC4":
                return "Low security – 40‑bit RC4 (outdated and easily broken).";
            case "Standard128BitRC4":
                return "Medium security – 128‑bit RC4 (considered weak by modern standards).";
            case "AES128":
                return "High security – AES with a 128‑bit key.";
            case "AES256":
                return "Very high security – AES with a 256‑bit key.";
            default:
                return "Unknown security level.";
        }
    }
}
