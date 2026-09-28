// Title: Save a password‑protected Aspose.Cells workbook to a MemoryStream and compare encrypted versus unencrypted stream sizes in C#
// AI Prompts: Write C# code that creates a Workbook, assigns Settings.Password, saves it to a MemoryStream using OoxmlSaveOptions, and prints the resulting stream length. | Add logic to save the same workbook without a password, then compare the lengths of the two MemoryStreams and output whether encryption added overhead. | Extend the example to use a different SaveFormat (e.g., Xls) while still applying password protection, and report the size difference between protected and unprotected streams.
// Common Searches: how to use Aspose.Cells to save an encrypted Excel file to a MemoryStream in C# | measure size increase caused by password protection with Aspose.Cells OoxmlSaveOptions | compare memory stream lengths of protected and unprotected workbooks using Aspose.Cells | C# Aspose.Cells encryption overhead when saving workbook to stream | retrieve byte length of password‑protected workbook saved to MemoryStream
// Tags: Aspose.Cells OoxmlSaveOptions password protect workbook in MemoryStream | C# compare encrypted and plain workbook stream lengths | encryption overhead measurement for Xlsx using Aspose.Cells | save workbook to MemoryStream with Settings.Password | detect size increase from password protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook, applying a password via Settings.Password, saving both unprotected and protected versions to MemoryStream (using OoxmlSaveOptions), and verifying that the protected stream length is larger due to encryption overhead.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Protect the workbook with a password (encryption)
            // Setting the password is sufficient for basic encryption
            string password = "SecretPassword";
            workbook.Settings.Password = password;

            // Save the unprotected workbook to a memory stream for baseline length
            using (MemoryStream unprotectedStream = new MemoryStream())
            {
                workbook.Save(unprotectedStream, SaveFormat.Xlsx);
                long unprotectedLength = unprotectedStream.Length;

                // Reset the stream position for reuse
                unprotectedStream.Position = 0;

                // Save the protected (encrypted) workbook to another memory stream
                using (MemoryStream protectedStream = new MemoryStream())
                {
                    // Use OoxmlSaveOptions to ensure encryption is applied
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                    workbook.Save(protectedStream, saveOptions);
                    long protectedLength = protectedStream.Length;

                    // Verify that the encrypted stream length reflects encryption overhead
                    Console.WriteLine($"Unprotected stream length: {unprotectedLength} bytes");
                    Console.WriteLine($"Protected (encrypted) stream length: {protectedLength} bytes");

                    if (protectedLength > unprotectedLength)
                    {
                        Console.WriteLine("Encryption overhead detected: protected stream is larger than unprotected stream.");
                    }
                    else
                    {
                        Console.WriteLine("No encryption overhead detected: check protection settings.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
