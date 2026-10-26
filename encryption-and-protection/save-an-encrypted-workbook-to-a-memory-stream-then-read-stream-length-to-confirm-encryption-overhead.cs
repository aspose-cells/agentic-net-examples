// Title: Encrypt an Excel workbook, save it to a MemoryStream, and read the encrypted file size with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# encrypt workbook and write to MemoryStream | Get size of password‑protected XLSX saved to a stream using Aspose.Cells | How to determine encryption overhead of an Excel file in .NET | Save encrypted workbook to MemoryStream and retrieve Length property with Aspose.Cells
// Tags: Aspose.Cells encrypt workbook to MemoryStream | password protected XLSX size calculation | Workbook.Settings.Password encryption example | SaveFormat.Xlsx encrypted output stream | MemoryStream length after Aspose.Cells save

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds data, applies a password via Workbook.Settings.Password to encrypt the file, saves the encrypted workbook to a MemoryStream in XLSX format, and then reads the MemoryStream.Length to obtain the size of the encrypted workbook, demonstrating how to measure encryption overhead.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");

        // Set a password to encrypt the workbook
        workbook.Settings.Password = "SecretPassword";

        // Save the encrypted workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);

            // Get the length of the stream (includes encryption overhead)
            long encryptedLength = ms.Length;

            Console.WriteLine($"Encrypted workbook size (bytes): {encryptedLength}");
        }
    }
}
