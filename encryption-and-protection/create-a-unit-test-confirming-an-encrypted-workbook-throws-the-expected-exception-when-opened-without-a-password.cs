// Title: C# Unit Test: Verify Encrypted Aspose.Cells Workbook Throws CellsException Without Password
// Description: Sample code that creates a temporary workbook, applies a password, saves it as XLSX, then attempts to load it without the password. The test asserts that Aspose.Cells throws a CellsException and cleans up the temporary file afterwards.
// Keywords: Aspose.Cells encrypted workbook unit test | C# CellsException password protected Excel | Aspose.Cells open encrypted file without password | Aspose.Cells .NET exception handling | unit test encrypted workbook Aspose | temporary file cleanup C# | XLSX password protection Aspose
// Common Searches: how to test opening a password protected workbook in Aspose.Cells | assert CellsException when loading encrypted Excel file without password | C# unit test for Aspose.Cells workbook encryption | Aspose.Cells exception for missing password | automated test encrypted workbook Aspose
// Developer Intent: Confirm that loading an encrypted Aspose.Cells workbook without supplying the correct password raises a CellsException.
// Use Cases: Continuous‑integration validation that password protection is enforced by Aspose.Cells. | Testing error‑handling paths for applications receiving encrypted Excel files. | Demonstrating proper creation and disposal of temporary encrypted files in automated tests.
// AI Prompts: Generate an MSTest/NUnit/xUnit test that creates an encrypted workbook with Aspose.Cells, saves it, then opens it without a password and asserts a CellsException is thrown. | Write C# code to capture the CellsException message when opening a password‑protected workbook without providing the password and verify its content. | Provide a reusable helper method that creates a temporary encrypted workbook, runs a test expecting an exception, and deletes the file afterward.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Sample code that creates a temporary workbook, applies a password, saves it as XLSX, then attempts to load it without the password. The test asserts that Aspose.Cells throws a CellsException and cleans up the temporary file afterwards.
    public class EncryptedWorkbookDemo
    {
        private const string Password = "Secret123";

        // Creates an encrypted workbook on disk and returns its path
        private static string CreateEncryptedWorkbook()
        {
            var wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Encrypted content");
            wb.Settings.Password = Password;

            string tempFile = Path.Combine(Path.GetTempPath(), $"Encrypted_{Guid.NewGuid()}.xlsx");
            wb.Save(tempFile, SaveFormat.Xlsx);
            wb.Dispose();

            return tempFile;
        }

        public static void Main()
        {
            string encryptedFile = null;

            try
            {
                // Arrange
                encryptedFile = CreateEncryptedWorkbook();

                // Ensure the file exists before attempting to load it
                if (!File.Exists(encryptedFile))
                    throw new FileNotFoundException("Encrypted workbook not found.", encryptedFile);

                // Act & Assert: loading without a password should throw CellsException
                try
                {
                    var wb = new Workbook(encryptedFile);
                    wb.Dispose(); // Should not reach here
                    Console.WriteLine("Test Failed: Workbook opened without password.");
                }
                catch (CellsException ex)
                {
                    // Expected outcome
                    Console.WriteLine($"Test Passed: Caught expected CellsException - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file
                if (!string.IsNullOrEmpty(encryptedFile) && File.Exists(encryptedFile))
                {
                    try
                    {
                        File.Delete(encryptedFile);
                    }
                    catch (Exception delEx)
                    {
                        Console.WriteLine($"Failed to delete temporary file: {delEx.Message}");
                    }
                }
            }
        }
    }
}
