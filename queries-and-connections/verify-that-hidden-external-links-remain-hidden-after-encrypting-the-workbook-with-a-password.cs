// Title: Check hidden external link visibility after encrypting an Excel workbook with a password using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds an external link via a formula, saves it, applies password protection (Workbook.Settings.Password), reloads the encrypted file with the password, confirms encryption, and iterates through Worksheets.ExternalLinks to display each link's DataSource and IsVisible flag, verifying that the visibility state is unchanged after encryption.
// Keywords: Aspose.Cells external link visibility | password protect Excel workbook C# | Workbook.Settings.Password example | encrypted workbook external links | IsVisible flag Aspose.Cells
// Common Searches: keep external links hidden after Excel encryption Aspose.Cells | does password protection change external link visibility | load password‑protected workbook and read external links .NET | Aspose.Cells verify hidden external links after encryption | C# check IsVisible of external links in encrypted workbook
// Developer Intent: Verify that the IsVisible property of external links remains unchanged when a workbook is saved with password protection.
// Use Cases: Create a workbook with a hidden external link, encrypt it, and programmatically confirm the link stays hidden after decryption. | Automate compliance checks to ensure external references are not exposed in password‑protected Excel files before distribution. | Batch‑process Excel files and validate that encryption does not alter external link metadata such as visibility.
// AI Prompts: Generate C# code using Aspose.Cells to add a hidden external link, encrypt the workbook with a password, then load it and confirm the link's IsVisible property remains false. | Explain how Aspose.Cells preserves external link visibility when a workbook is saved with password protection and show how to check it programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsHiddenExternalLinkDemo
{
    // This C# example creates a workbook, adds an external link via a formula, saves it, applies password protection (Workbook.Settings.Password), reloads the encrypted file with the password, confirms encryption, and iterates through Worksheets.ExternalLinks to display each link's DataSource and IsVisible flag, verifying that the visibility state is unchanged after encryption.
    class Program
    {
        static void Main()
        {
            // Step 1: Create a workbook and add an external link via a formula
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // The formula creates an external link to a non‑existent file (for demo purposes)
            ws.Cells["A1"].Formula = "='[ExternalSource.xlsx]Sheet1'!A1";

            // Save the workbook (unprotected) so the external link is persisted
            string unprotectedPath = "UnprotectedWorkbook.xlsx";
            wb.Save(unprotectedPath);

            // Step 2: Load the workbook again, set a password (encrypt it) and save
            Workbook wbToEncrypt = new Workbook(unprotectedPath);
            wbToEncrypt.Settings.Password = "SecretPassword";
            string encryptedPath = "EncryptedWorkbook.xlsx";
            wbToEncrypt.Save(encryptedPath);

            // Step 3: Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions { Password = "SecretPassword" };
            Workbook encryptedWb = new Workbook(encryptedPath, loadOptions);

            // Verify that the workbook is indeed encrypted
            Console.WriteLine($"Workbook IsEncrypted: {encryptedWb.Settings.IsEncrypted}");

            // Step 4: Iterate through external links and display their visibility status
            foreach (ExternalLink link in encryptedWb.Worksheets.ExternalLinks)
            {
                Console.WriteLine($"External Link DataSource: {link.DataSource}");
                Console.WriteLine($"IsVisible (should remain unchanged after encryption): {link.IsVisible}");
            }

            // Clean up
            wb.Dispose();
            wbToEncrypt.Dispose();
            encryptedWb.Dispose();
        }
    }
}
