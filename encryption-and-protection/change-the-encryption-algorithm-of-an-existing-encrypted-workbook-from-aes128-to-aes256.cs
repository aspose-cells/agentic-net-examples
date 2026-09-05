// Title: Re‑encrypt an AES‑128 protected Excel workbook to AES‑256 with Aspose.Cells for .NET
// AI Prompts: Load an AES‑128 encrypted .xlsx file using a password, change its encryption setting to AES‑256, and save it as a new file with Aspose.Cells in C#. | Open a password‑protected workbook, upgrade the encryption algorithm to AES‑256, and write the updated workbook to a different path using the Aspose.Cells API. | Programmatically replace AES‑128 encryption with AES‑256 for an existing Excel file while preserving the original password in C#.
// Common Searches: aspnet change Excel file encryption from AES-128 to AES-256 using Aspose.Cells | how to upgrade password protected XLSX encryption to AES-256 in C# | Aspose.Cells re‑encrypt workbook with stronger AES algorithm | convert AES-128 encrypted workbook to AES-256 programmatically | C# load encrypted XLSX and save with AES-256 Aspose.Cells
// Tags: Aspose.Cells AES-256 re-encryption | C# upgrade Excel encryption algorithm Aspose.Cells | load password protected XLSX Aspose.Cells | save workbook with stronger encryption Aspose.Cells | change workbook encryption from AES-128 to AES-256

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // // Loads an AES‑128 encrypted Excel workbook, sets the same password (defaulting to AES‑256 encryption), and saves the workbook as a new file with the stronger AES‑256 protection.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "EncryptedWorkbook_AES128.xlsx";
            const string outputFile = "EncryptedWorkbook_AES256.xlsx";
            const string password = "yourPassword";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the existing encrypted workbook (AES‑128) using its password
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                var workbook = new Workbook(inputFile, loadOptions);

                // Set the password for the new workbook (default encryption is AES‑256)
                workbook.Settings.Password = password;

                // Save the workbook with the new encryption
                workbook.Save(outputFile, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook successfully re‑encrypted and saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
