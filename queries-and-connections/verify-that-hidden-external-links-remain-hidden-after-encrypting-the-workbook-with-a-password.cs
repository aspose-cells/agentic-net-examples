using System;
using Aspose.Cells;

namespace HiddenExternalLinkVerification
{
    class Program
    {
        static void Main()
        {
            // ------------------- Create workbook with hidden external link -------------------
            Workbook workbook = new Workbook();                         // create new workbook
            Worksheet sheet = workbook.Worksheets[0];                  // get first worksheet
            sheet.Cells["A1"].PutValue("Demo");                        // add some data

            // Add an external link; by default such links are hidden (IsVisible == false)
            workbook.Worksheets.ExternalLinks.Add("external.xlsx", new string[] { "Sheet1!A1" });

            // Retrieve the external link and check its visibility before encryption
            ExternalLink linkBefore = workbook.Worksheets.ExternalLinks[0];
            bool isVisibleBefore = linkBefore.IsVisible;
            Console.WriteLine($"External link visibility before encryption: {isVisibleBefore}");

            // ------------------- Encrypt workbook with password -------------------
            workbook.Settings.Password = "SecretPwd123";               // set encryption password
            workbook.Save("HiddenLinkEncrypted.xlsx");                 // save encrypted file

            // ------------------- Load encrypted workbook and verify link visibility -------------------
            LoadOptions loadOptions = new LoadOptions { Password = "SecretPwd123" };
            Workbook loadedWorkbook = new Workbook("HiddenLinkEncrypted.xlsx", loadOptions);

            // Access the external link after loading
            ExternalLink linkAfter = loadedWorkbook.Worksheets.ExternalLinks[0];
            bool isVisibleAfter = linkAfter.IsVisible;
            Console.WriteLine($"External link visibility after encryption: {isVisibleAfter}");

            // Final verification output
            if (isVisibleBefore == isVisibleAfter)
                Console.WriteLine("Hidden external link remained hidden after encryption.");
            else
                Console.WriteLine("External link visibility changed after encryption.");
        }
    }
}