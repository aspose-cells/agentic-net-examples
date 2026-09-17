// Title: C# unit test to verify that opening an encrypted Aspose.Cells workbook without a password throws CellsException
// AI Prompts: Write an MSTest method that creates a password‑protected workbook, saves it to a MemoryStream, and asserts that constructing a new Workbook from the stream without a password raises a CellsException. | Generate an xUnit test that encrypts an XLSX file using Workbook.Settings.Password, then attempts to load it without credentials and checks for the expected exception type.
// Common Searches: aspocells test opening password protected xlsx without password throws exception | how to assert error when loading encrypted Excel workbook in C# using Aspose.Cells | unit testing Aspose.Cells encrypted workbook loading failure | C# MSTest for verifying missing password error on encrypted workbook
// Tags: Aspose.Cells encrypted workbook error handling | C# unit testing password protected Excel | Workbook.Settings.Password usage in tests | verify encrypted XLSX load failure | assert error on workbook open without password

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // // Demonstrates creating a simple workbook, applying a password via Workbook.Settings.Password, saving it to a MemoryStream, and then attempting to load it without providing the password. The code catches the expected CellsException, confirming proper error handling for missing passwords.
    class Program
    {
        static void Main()
        {
            OpenEncryptedWorkbookWithoutPassword_ShouldThrowException();
        }

        static void OpenEncryptedWorkbookWithoutPassword_ShouldThrowException()
        {
            try
            {
                // Create a simple workbook with one worksheet and some data
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Test");

                // Encrypt the workbook with a password
                workbook.Settings.Password = "SecretPassword";

                // Save the encrypted workbook to a memory stream
                using (var encryptedStream = new MemoryStream())
                {
                    workbook.Save(encryptedStream, SaveFormat.Xlsx);
                    encryptedStream.Position = 0; // Reset stream for reading

                    // Attempt to load the encrypted workbook without providing a password
                    try
                    {
                        var loadedWorkbook = new Workbook(encryptedStream);
                        Console.WriteLine("Error: Workbook opened without password; an exception was expected.");
                    }
                    catch (CellsException ex)
                    {
                        // Expected exception for missing password
                        Console.WriteLine($"Expected exception caught: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
