using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    public class EncryptedWorkbookTests
    {
        private const string Password = "Secret123";

        // Helper to create a temporary encrypted workbook file
        private string CreateEncryptedWorkbook()
        {
            // Generate a unique file name with .xlsx extension
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            try
            {
                // Create a new workbook and set its encryption password
                Workbook wb = new Workbook();
                wb.Settings.Password = Password; // encrypt the workbook

                // Save the encrypted workbook to the temporary location
                wb.Save(tempFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating encrypted workbook: {ex.Message}");
                throw;
            }

            return tempFile;
        }

        // Simulates the test: opening an encrypted workbook without providing a password should throw CellsException
        public void OpeningEncryptedWorkbookWithoutPassword_ShouldThrowException()
        {
            // Arrange: create an encrypted workbook file
            string encryptedFile = CreateEncryptedWorkbook();

            try
            {
                // Act: attempt to open the encrypted workbook without a password
                try
                {
                    if (!File.Exists(encryptedFile))
                        throw new FileNotFoundException("Encrypted workbook file not found.", encryptedFile);

                    // This should throw a CellsException because the password is missing
                    Workbook wb = new Workbook(encryptedFile);
                    Console.WriteLine("ERROR: Workbook opened without password – test failed.");
                }
                catch (CellsException)
                {
                    // Expected path: exception thrown due to missing password
                    Console.WriteLine("PASS: CellsException thrown as expected when opening encrypted workbook without password.");
                }
                catch (Exception ex)
                {
                    // Any other exception means the test did not behave as expected
                    Console.WriteLine($"FAIL: Unexpected exception type: {ex.GetType().Name} - {ex.Message}");
                }
            }
            finally
            {
                // Clean up the temporary file
                try
                {
                    if (File.Exists(encryptedFile))
                        File.Delete(encryptedFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not delete temporary file. {ex.Message}");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                var tests = new EncryptedWorkbookTests();
                tests.OpeningEncryptedWorkbookWithoutPassword_ShouldThrowException();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}