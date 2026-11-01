// Title: Write a C# unit test that confirms a CellsException is thrown when loading an encrypted Aspose.Cells workbook without providing a password
// Common Searches: how to assert CellsException for password‑protected Excel workbook in Aspose.Cells C# unit test | C# unit test for loading encrypted .xlsx with Aspose.Cells without password | verify exception message contains password when opening encrypted workbook using Aspose.Cells | Aspose.Cells C# test expecting error when opening protected workbook without password | unit testing Aspose.Cells workbook encryption handling in .NET
// Tags: Aspose.Cells encrypted workbook unit test | CellsException handling for password protected Excel | load encrypted .xlsx without password Aspose.Cells | C# test for workbook password validation | assert exception message contains password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a simple Workbook, applies a password via Settings.Password, saves it to a MemoryStream as XLSX, and then attempts to load the encrypted stream without supplying a password. The code catches the resulting CellsException and checks that the exception message includes the word "password", demonstrating how to write a C# unit test that validates proper error handling for password‑protected workbooks in Aspose.Cells.
    public class EncryptedWorkbookDemo
    {
        public static void Main()
        {
            try
            {
                // Create a simple workbook and encrypt it with a password.
                var originalWorkbook = new Workbook();
                var sheet = originalWorkbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Secret Data");
                originalWorkbook.Settings.Password = "MySecretPwd";

                // Save the encrypted workbook to a memory stream.
                using (var encryptedStream = new MemoryStream())
                {
                    originalWorkbook.Save(encryptedStream, SaveFormat.Xlsx);
                    encryptedStream.Position = 0; // Reset stream position for reading.

                    // Attempt to load the encrypted workbook without providing a password.
                    try
                    {
                        var wb = new Workbook(encryptedStream);
                        Console.WriteLine("Unexpectedly opened the encrypted workbook without a password.");
                    }
                    catch (CellsException ex)
                    {
                        // Verify that the exception message indicates a password is required.
                        if (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine("Caught expected CellsException: " + ex.Message);
                        }
                        else
                        {
                            Console.WriteLine("Caught CellsException, but message does not mention password: " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Catch any unexpected exceptions.
                Console.WriteLine("Unexpected error: " + e);
            }
        }
    }
}
