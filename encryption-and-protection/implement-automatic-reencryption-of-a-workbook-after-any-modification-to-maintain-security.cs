// Title: C# – Auto‑re‑encrypt an Aspose.Cells workbook after each modification
// Description: Demonstrates how to create a password‑protected workbook, save it, load it with LoadOptions, modify cells, re‑apply the password (or stronger encryption via SetEncryptionOptions), and verify that the file remains encrypted—all in C# using Aspose.Cells.
// Keywords: Aspose.Cells auto re‑encrypt workbook | C# workbook password encryption | re‑apply password after edit Aspose.Cells | SetEncryptionOptions Aspose.Cells .NET | load encrypted Excel modify save | programmatic Excel encryption C# | secure Aspose.Cells workbook
// Common Searches: how to re‑encrypt an Aspose.Cells workbook after changes | Aspose.Cells .NET update encrypted workbook | C# set stronger encryption for Excel file with Aspose | verify password protection after saving Aspose.Cells workbook | auto‑re‑encrypt Excel file using Aspose.Cells
// Developer Intent: Automatically re‑apply password protection to an Aspose.Cells workbook whenever its content is altered.
// Use Cases: Create a new workbook, protect it with a password, and store it securely. | Open an existing encrypted workbook, edit data, and save it while preserving or upgrading the encryption. | Confirm that a re‑saved workbook still requires the password to open.
// AI Prompts: Show C# code that automatically re‑encrypts an Aspose.Cells workbook after each cell update, including optional stronger encryption settings. | Provide an example of loading an encrypted Excel file with Aspose.Cells, modifying it, re‑applying the password, and verifying the protection. | Explain how to detect workbook changes in Aspose.Cells and trigger re‑encryption programmatically.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSecurityDemo
{
    // Demonstrates how to create a password‑protected workbook, save it, load it with LoadOptions, modify cells, re‑apply the password (or stronger encryption via SetEncryptionOptions), and verify that the file remains encrypted—all in C# using Aspose.Cells.
    public class AutomaticReEncryptionDemo
    {
        public static void Run()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and set an initial password
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Original data");

                // Set password to encrypt the workbook
                string password = "SecurePwd123";
                workbook.Settings.Password = password;                 // encrypt

                // Save the encrypted workbook
                string encryptedPath = "EncryptedWorkbook.xlsx";
                workbook.Save(encryptedPath);                           // save

                // -----------------------------------------------------------------
                // 2. Load the encrypted workbook, modify it, and re‑encrypt
                // -----------------------------------------------------------------
                if (!File.Exists(encryptedPath))
                    throw new FileNotFoundException($"File not found: {encryptedPath}");

                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password                                 // load with password
                };
                Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions); // load

                // Perform some modifications
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                loadedSheet.Cells["B2"].PutValue("Modified after load");

                // Re‑apply encryption after modification
                // (re‑setting the password forces re‑encryption)
                loadedWorkbook.Settings.Password = password;

                // Optionally set stronger encryption options
                loadedWorkbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // Save the re‑encrypted workbook
                string reEncryptedPath = "ReEncryptedWorkbook.xlsx";
                loadedWorkbook.Save(reEncryptedPath);                  // save

                // -----------------------------------------------------------------
                // 3. Verify that the workbook is still encrypted
                // -----------------------------------------------------------------
                if (!File.Exists(reEncryptedPath))
                    throw new FileNotFoundException($"File not found: {reEncryptedPath}");

                LoadOptions verifyOptions = new LoadOptions { Password = password };
                Workbook verifyWorkbook = new Workbook(reEncryptedPath, verifyOptions);
                Console.WriteLine("Verification cell value: " +
                    verifyWorkbook.Worksheets[0].Cells["B2"].Value);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutomaticReEncryptionDemo.Run();
        }
    }
}
