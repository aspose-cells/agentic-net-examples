// Title: Measure the Size Increase of an Encrypted vs. Unencrypted Workbook with Aspose.Cells for .NET
// Description: The sample creates a workbook, saves it as an unencrypted XLSX file, records its size, applies a password via Workbook.Settings.Password, saves the encrypted file, records the new size, and uses FileFormatUtil.DetectFileFormat to confirm encryption status while displaying the size difference.
// Keywords: Aspose.Cells | C# workbook encryption | Excel file size overhead | password protected XLSX size | FileFormatUtil DetectFileFormat | .NET encryption impact | Excel storage increase encryption
// Common Searches: Aspose.Cells compare encrypted and unencrypted file size | C# get size of password protected Excel file | How much does Excel encryption add to file size | Detect if XLSX is encrypted using Aspose.Cells | Measure storage overhead of workbook password protection
// Developer Intent: Find out how much additional storage an Excel workbook consumes after applying password protection with Aspose.Cells.
// Use Cases: Report encryption overhead for regulatory compliance | Ensure encrypted workbooks fit within cloud storage limits | Batch‑process multiple files to decide if password protection is viable | Validate that encryption does not exceed size thresholds for email attachments
// AI Prompts: Write C# code that creates a workbook, saves it unencrypted, encrypts it with a password, and prints the original and encrypted file sizes using Aspose.Cells. | Provide a function that takes a file path and password, encrypts the workbook, and returns the size delta. | Explain how FileFormatUtil.DetectFileFormat can be used to verify whether an Excel file is encrypted after saving.

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptionImpactDemo
{
    // The sample creates a workbook, saves it as an unencrypted XLSX file, records its size, applies a password via Workbook.Settings.Password, saves the encrypted file, records the new size, and uses FileFormatUtil.DetectFileFormat to confirm encryption status while displaying the size difference.
    class Program
    {
        static void Main()
        {
            // Paths for the files
            string unencryptedPath = "UnencryptedWorkbook.xlsx";
            string encryptedPath = "EncryptedWorkbook.xlsx";

            // -------------------------------------------------
            // Create a sample workbook and add some data
            // -------------------------------------------------
            Workbook wb = new Workbook(); // create new workbook
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for encryption impact test.");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345);

            // -------------------------------------------------
            // Save the unencrypted workbook
            // -------------------------------------------------
            wb.Save(unencryptedPath);
            long unencryptedSize = new FileInfo(unencryptedPath).Length;

            // -------------------------------------------------
            // Apply password protection (encryption)
            // -------------------------------------------------
            wb.Settings.Password = "StrongPassword123";
            // Optionally set encryption options (default is sufficient for most cases)
            // wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            wb.Save(encryptedPath);
            long encryptedSize = new FileInfo(encryptedPath).Length;

            // -------------------------------------------------
            // Verify encryption status using FileFormatInfo
            // -------------------------------------------------
            FileFormatInfo unencInfo = FileFormatUtil.DetectFileFormat(unencryptedPath);
            FileFormatInfo encInfo = FileFormatUtil.DetectFileFormat(encryptedPath);

            Console.WriteLine($"Unencrypted file size: {unencryptedSize} bytes");
            Console.WriteLine($"Encrypted file size:   {encryptedSize} bytes");
            Console.WriteLine($"Size increase:        {encryptedSize - unencryptedSize} bytes");
            Console.WriteLine();
            Console.WriteLine($"Is unencrypted file encrypted? {unencInfo.IsEncrypted}");
            Console.WriteLine($"Is encrypted file encrypted?   {encInfo.IsEncrypted}");

            // Clean up
            wb.Dispose();
        }
    }
}
