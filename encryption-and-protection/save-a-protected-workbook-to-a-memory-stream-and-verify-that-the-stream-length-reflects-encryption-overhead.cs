// Title: Save a password‑protected Excel workbook with Aspose.Cells to a MemoryStream and measure encryption overhead in C#
// Common Searches: C# Aspose.Cells save encrypted XLSX to MemoryStream and get stream length | How to determine size difference between password protected and unprotected Excel files using Aspose.Cells | Measure encryption overhead when saving a workbook with a password in Aspose.Cells C# | Retrieve MemoryStream length after saving a protected workbook with Aspose.Cells | Compare protected and unprotected workbook stream sizes in .NET
// Tags: Aspose.Cells password encryption to MemoryStream | measure encrypted XLSX stream length | compare protected vs unprotected workbook size | C# detect Excel encryption overhead | save workbook as encrypted stream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills a few cells, applies a password to encrypt it, and saves the encrypted workbook to a MemoryStream. It then creates an identical unprotected workbook, saves it to another MemoryStream, and writes both stream lengths to the console, demonstrating the size increase caused by encryption.
class Program
{
    static void Main()
    {
        // ---------- Create a workbook ----------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B1"].PutValue(42);
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // ---------- Protect (encrypt) the workbook ----------
        // Setting a password encrypts the file when saved.
        workbook.Settings.Password = "MySecretPwd";

        // ---------- Save protected workbook to a memory stream ----------
        using (MemoryStream protectedStream = new MemoryStream())
        {
            workbook.Save(protectedStream, SaveFormat.Xlsx);
            long protectedLength = protectedStream.Length;

            // ---------- Save an unprotected version for comparison ----------
            Workbook unprotectedWb = new Workbook();
            Worksheet unprotectedSheet = unprotectedWb.Worksheets[0];
            unprotectedSheet.Cells["A1"].PutValue("Sample Text");
            unprotectedSheet.Cells["B1"].PutValue(42);
            unprotectedSheet.Cells["A2"].PutValue(DateTime.Now);

            using (MemoryStream unprotectedStream = new MemoryStream())
            {
                unprotectedWb.Save(unprotectedStream, SaveFormat.Xlsx);
                long unprotectedLength = unprotectedStream.Length;

                // ---------- Verify encryption overhead ----------
                Console.WriteLine($"Protected stream length: {protectedLength}");
                Console.WriteLine($"Unprotected stream length: {unprotectedLength}");
                Console.WriteLine($"Encryption overhead detected: {protectedLength > unprotectedLength}");
            }
        }
    }
}
