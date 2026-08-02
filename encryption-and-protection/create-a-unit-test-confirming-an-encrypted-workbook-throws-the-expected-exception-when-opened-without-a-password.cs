// Title: C# unit test: Verify CellsException when opening an encrypted Aspose.Cells workbook without a password
// Description: Creates a temporary encrypted Excel file by setting Workbook.Settings.Password, then runs two scenarios: (1) loads the file with default LoadOptions and asserts that a CellsException containing the word "password" is thrown, and (2) loads the same file with the correct password to confirm successful opening and that the workbook remains encrypted. Includes SetUp and TearDown for file management.
// Keywords: Aspose.Cells | C# | .NET | encrypted workbook | password protection | CellsException | LoadOptions | unit test | MSTest | NUnit | XUnit | Excel | exception handling | test automation
// Common Searches: Aspose.Cells unit test encrypted workbook | assert exception opening password protected Excel with Aspose.Cells | CellsException missing password .NET | how to test encrypted workbook in C# | load encrypted Excel without password Aspose.Cells | unit test for workbook encryption Aspose.Cells
// Developer Intent: Ensure that attempting to open an Aspose.Cells workbook encrypted with a password, without providing that password, raises the expected CellsException.
// Use Cases: Generate a temporary encrypted workbook in test setup to isolate test data. | Attempt to load the encrypted file without a password and assert that a CellsException mentioning "password" is thrown. | Load the same file with the correct password and verify the workbook opens and reports IsEncrypted = true. | Automatically delete the temporary file after the test run.
// AI Prompts: Create an MSTest method that encrypts a workbook with Aspose.Cells, then asserts a CellsException is thrown when opening it without a password. | Write an XUnit test that verifies the exception message includes the word "password" when a protected workbook is loaded without credentials. | Provide a NUnit example that loads an encrypted Excel file using LoadOptions with and without the password and checks the appropriate outcomes.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a temporary encrypted Excel file by setting Workbook.Settings.Password, then runs two scenarios: (1) loads the file with default LoadOptions and asserts that a CellsException containing the word "password" is thrown, and (2) loads the same file with the correct password to confirm successful opening and that the workbook remains encrypted. Includes SetUp and TearDown for file management.
    public class EncryptedWorkbookDemo
    {
        private const string Password = "Secret123!";
        private string _encryptedFilePath;

        // Create a temporary encrypted workbook
        public void SetUp()
        {
            try
            {
                _encryptedFilePath = Path.Combine(Path.GetTempPath(), $"Encrypted_{Guid.NewGuid()}.xlsx");

                var workbook = new Workbook();
                workbook.Settings.Password = Password; // encrypt the workbook
                workbook.Save(_encryptedFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SetUp failed: {ex.Message}");
                throw;
            }
        }

        // Delete the temporary file
        public void TearDown()
        {
            try
            {
                if (!string.IsNullOrEmpty(_encryptedFilePath) && File.Exists(_encryptedFilePath))
                {
                    File.Delete(_encryptedFilePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TearDown failed: {ex.Message}");
            }
        }

        // Test opening without providing a password – should throw CellsException
        public void TestOpeningEncryptedWorkbookWithoutPassword()
        {
            try
            {
                var loadOptions = new LoadOptions(); // no password set

                // Ensure the file exists before attempting to load
                if (!File.Exists(_encryptedFilePath))
                    throw new FileNotFoundException("Encrypted workbook not found.", _encryptedFilePath);

                // This line is expected to throw
                var wb = new Workbook(_encryptedFilePath, loadOptions);
                Console.WriteLine("ERROR: Expected exception was not thrown.");
            }
            catch (CellsException ex)
            {
                // Verify that the exception message mentions password
                if (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("PASS: Correct exception thrown for missing password.");
                }
                else
                {
                    Console.WriteLine($"FAIL: Exception thrown but message does not mention password. Message: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FAIL: Unexpected exception type: {ex.GetType().Name}, Message: {ex.Message}");
            }
        }

        // Test opening with the correct password – should succeed
        public void TestOpeningEncryptedWorkbookWithCorrectPassword()
        {
            try
            {
                var loadOptions = new LoadOptions { Password = Password };

                if (!File.Exists(_encryptedFilePath))
                    throw new FileNotFoundException("Encrypted workbook not found.", _encryptedFilePath);

                var wb = new Workbook(_encryptedFilePath, loadOptions);

                if (wb.Settings.IsEncrypted)
                {
                    Console.WriteLine("PASS: Workbook opened successfully with correct password and is encrypted.");
                }
                else
                {
                    Console.WriteLine("FAIL: Workbook opened but IsEncrypted flag is false.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FAIL: Exception while opening with correct password: {ex.GetType().Name}, Message: {ex.Message}");
            }
        }

        // Entry point to run the demo
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
        }
    }
}
