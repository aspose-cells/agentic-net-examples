// Title: C# Unit Test: Verify CellsException is thrown when opening an encrypted Aspose.Cells workbook without a password
// Description: Creates a temporary workbook, applies a password via Workbook.Settings.Password, saves it, then attempts to load the file with LoadOptions that omit the password. The test asserts that a CellsException containing the word "Password" is raised. A second check loads the same file with the correct password to confirm successful opening and that the workbook reports as encrypted. The temporary file is deleted after the test.
// Keywords: Aspose.Cells encrypted workbook unit test | CellsException password missing | LoadOptions without password C# | Aspose.Cells password protection test | C# Excel encryption unit testing
// Common Searches: Aspose.Cells unit test for password protected workbook | How to assert exception when opening encrypted Excel with Aspose.Cells | C# test loading encrypted workbook without password | Validate password requirement in Aspose.Cells LoadOptions | MSTest/NUnit example for Aspose.Cells encryption
// Developer Intent: Write a C# unit test that confirms opening a password‑protected workbook without supplying the password triggers the expected CellsException.
// Use Cases: Ensure that a workbook saved with Settings.Password cannot be opened without the same password. | Automatically verify that the exception message clearly indicates a missing password for security compliance. | Demonstrate correct handling of both failure (no password) and success (correct password) scenarios in automated test suites.
// AI Prompts: Generate an MSTest method that creates an encrypted workbook, attempts to open it without a password, and asserts that a CellsException containing "Password" is thrown. | Provide a NUnit test case for Aspose.Cells that checks failure without a password and success with the correct password, including proper temporary file cleanup.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a temporary workbook, applies a password via Workbook.Settings.Password, saves it, then attempts to load the file with LoadOptions that omit the password. The test asserts that a CellsException containing the word "Password" is raised. A second check loads the same file with the correct password to confirm successful opening and that the workbook reports as encrypted. The temporary file is deleted after the test.
    public class EncryptedWorkbookDemo
    {
        private const string Password = "Secret123!";
        private string _encryptedFilePath;

        // Create a new encrypted workbook and save it to a temporary file
        private void SetUp()
        {
            try
            {
                var workbook = new Workbook();
                workbook.Settings.Password = Password;

                _encryptedFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(_encryptedFilePath);
                Console.WriteLine($"Encrypted workbook created at: {_encryptedFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SetUp failed: {ex.Message}");
                throw;
            }
        }

        // Delete the temporary file if it exists
        private void TearDown()
        {
            try
            {
                if (!string.IsNullOrEmpty(_encryptedFilePath) && File.Exists(_encryptedFilePath))
                {
                    File.Delete(_encryptedFilePath);
                    Console.WriteLine("Temporary file deleted.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TearDown failed: {ex.Message}");
            }
        }

        // Test loading the encrypted workbook without providing a password
        private void TestOpeningEncryptedWorkbookWithoutPassword()
        {
            Console.WriteLine("Running TestOpeningEncryptedWorkbookWithoutPassword...");

            try
            {
                var loadOptions = new LoadOptions(); // No password set
                var wb = new Workbook(_encryptedFilePath, loadOptions);
                Console.WriteLine("FAIL: Workbook loaded without password (expected exception).");
            }
            catch (CellsException ex)
            {
                // Expected exception – verify it mentions password requirement
                if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("PASS: Expected exception thrown – " + ex.Message);
                }
                else
                {
                    Console.WriteLine($"FAIL: Exception thrown but message does not indicate password issue: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FAIL: Unexpected exception type: {ex.GetType().Name} - {ex.Message}");
            }
        }

        // Test loading the encrypted workbook with the correct password
        private void TestOpeningEncryptedWorkbookWithCorrectPassword()
        {
            Console.WriteLine("Running TestOpeningEncryptedWorkbookWithCorrectPassword...");

            try
            {
                var loadOptions = new LoadOptions { Password = Password };
                var wb = new Workbook(_encryptedFilePath, loadOptions);

                if (wb.Settings.IsEncrypted)
                {
                    Console.WriteLine("PASS: Workbook loaded with correct password and reports as encrypted.");
                }
                else
                {
                    Console.WriteLine("FAIL: Workbook loaded but does not report as encrypted.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FAIL: Exception while loading with correct password: {ex.GetType().Name} - {ex.Message}");
            }
        }

        // Entry point
        public static void Main()
        {
            var demo = new EncryptedWorkbookDemo();

            try
            {
                demo.SetUp();

                demo.TestOpeningEncryptedWorkbookWithoutPassword();
                demo.TestOpeningEncryptedWorkbookWithCorrectPassword();
            }
            finally
            {
                demo.TearDown();
            }

            Console.WriteLine("Demo completed.");
        }
    }
}
