using System;
using Aspose.Cells;

namespace AsposeCellsHiddenExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook and add an external link
            Workbook wb = new Workbook();
            // Add an external link that points to a dummy source file and a cell reference
            wb.Worksheets.ExternalLinks.Add("external_source.xlsx", new string[] { "Sheet1!A1" });

            // Retrieve the added external link
            ExternalLink link = wb.Worksheets.ExternalLinks[0];

            // Store the initial visibility state (should be false for hidden links)
            bool initialVisibility = link.IsVisible;
            Console.WriteLine($"Initial external link visibility: {initialVisibility}");

            // Step 2: Encrypt the workbook with a password
            wb.Settings.Password = "SecretPwd123";
            string encryptedFile = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedFile);

            // Step 3: Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = "SecretPwd123";
            Workbook loadedWb = new Workbook(encryptedFile, loadOptions);

            // Verify that the workbook reports being encrypted
            Console.WriteLine($"Workbook IsEncrypted after load: {loadedWb.Settings.IsEncrypted}");

            // Step 4: Check the external link visibility after decryption
            ExternalLink loadedLink = loadedWb.Worksheets.ExternalLinks[0];
            bool postLoadVisibility = loadedLink.IsVisible;
            Console.WriteLine($"External link visibility after encryption/decryption: {postLoadVisibility}");

            // Confirm that the visibility state has not changed
            Console.WriteLine($"Visibility unchanged: {initialVisibility == postLoadVisibility}");
        }
    }
}