// Title: Save a password‑protected Aspose.Cells workbook to a MemoryStream and compare its size to detect encryption overhead (C#)
// Common Searches: C# how to encrypt an Excel workbook with Aspose.Cells and write it to a MemoryStream | measure size difference between protected and unprotected XLSX using Aspose.Cells in .NET | Aspose.Cells protect workbook with password and get memory stream length | detect encryption overhead when saving a protected workbook to MemoryStream in C# | compare unencrypted and encrypted workbook stream sizes with Aspose.Cells
// Tags: Aspose.Cells workbook password protection XLSX | saving Excel workbook to MemoryStream with Aspose.Cells | detecting encryption overhead in protected XLSX files | C# stream length comparison for encrypted Excel | Aspose.Cells encryption size impact analysis

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds data, saves it unprotected to a MemoryStream, records its length, then applies password protection with ProtectionType.All, saves the protected workbook to another MemoryStream, records the new length, and checks whether the protected stream is larger, indicating encryption overhead.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Save the unprotected workbook to a memory stream
        using (MemoryStream unprotectedStream = new MemoryStream())
        {
            wb.Save(unprotectedStream, SaveFormat.Xlsx);
            long unprotectedLength = unprotectedStream.Length;
            Console.WriteLine($"Unprotected stream length: {unprotectedLength} bytes");

            // Protect the workbook with a password (encryption)
            wb.Protect(ProtectionType.All, "MySecretPassword");

            // Save the protected workbook to another memory stream
            using (MemoryStream protectedStream = new MemoryStream())
            {
                wb.Save(protectedStream, SaveFormat.Xlsx);
                long protectedLength = protectedStream.Length;
                Console.WriteLine($"Protected stream length: {protectedLength} bytes");

                // Verify that the protected stream length reflects encryption overhead
                if (protectedLength > unprotectedLength)
                {
                    Console.WriteLine("Encryption overhead detected: protected stream is larger than unprotected stream.");
                }
                else
                {
                    Console.WriteLine("No encryption overhead detected (unexpected).");
                }
            }
        }
    }
}
