using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyHiddenExternalLinksAfterEncryption
    {
        public static void Run()
        {
            // Path for the encrypted workbook
            const string encryptedPath = "HiddenExternalLink_Encrypted.xlsx";

            // Ensure any previous file is removed to avoid load conflicts
            if (File.Exists(encryptedPath))
            {
                File.Delete(encryptedPath);
            }

            Workbook wb = null;
            Workbook loadedWb = null;

            try
            {
                // -------------------- Create workbook with external link --------------------
                wb = new Workbook(); // create new workbook

                // Add an external link to a non‑existent file (the file does not need to exist for the link object)
                wb.Worksheets.ExternalLinks.Add("external_source.xlsx", new[] { "Sheet1!A1" });

                // Retrieve the external link
                ExternalLink link = wb.Worksheets.ExternalLinks[0];

                // Check visibility before encryption (IsVisible is read‑only; hidden links return false)
                Console.WriteLine($"Before encryption - External link visible: {link.IsVisible}");

                // -------------------- Encrypt workbook with password --------------------
                wb.Settings.Password = "SecretPwd123"; // set encryption password
                wb.Save(encryptedPath); // save encrypted workbook

                // Verify workbook reports as encrypted
                Console.WriteLine($"Workbook IsEncrypted after save: {wb.Settings.IsEncrypted}");

                // -------------------- Load encrypted workbook --------------------
                LoadOptions loadOpts = new LoadOptions { Password = "SecretPwd123" };
                loadedWb = new Workbook(encryptedPath, loadOpts); // load with password

                // Retrieve the external link from the loaded workbook
                ExternalLink loadedLink = loadedWb.Worksheets.ExternalLinks[0];

                // Verify that the hidden status is preserved after encryption/decryption
                Console.WriteLine($"After decryption - External link visible: {loadedLink.IsVisible}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up resources
                wb?.Dispose();
                loadedWb?.Dispose();
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            VerifyHiddenExternalLinksAfterEncryption.Run();
        }
    }
}