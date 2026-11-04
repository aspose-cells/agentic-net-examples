// Title: Write a C# unit test that confirms a CellsException is thrown when loading an Aspose.Cells workbook encrypted with a password without supplying the password
// Common Searches: C# unit test for Aspose.Cells encrypted workbook loading without password | How to assert CellsException for invalid password in Aspose.Cells | Testing password protection error handling in Aspose.Cells using NUnit | Verify exception when opening password‑protected Excel file with Aspose.Cells in .NET | Aspose.Cells load encrypted workbook without password throws exception
// Tags: Aspose.Cells password protection unit testing | CellsException invalid password handling | encrypted workbook loading without password .NET | temporary file cleanup in Aspose.Cells tests | NUnit Aspose.Cells workbook encryption test

using System;
using System.IO;
using Aspose.Cells;

namespace EncryptedWorkbookDemo
{
    // The example creates a workbook, adds data, encrypts it with a password via workbook.Settings.Password, saves it to a temporary file, then attempts to load the file without providing a password. It catches the expected CellsException containing "Invalid password", flags the test as passed, reports failure if the exception is not thrown, and finally deletes the temporary file.
    class Program
    {
        static void Main()
        {
            string tempFilePath = null;

            try
            {
                // ----- Arrange -----
                // Create a new workbook and add some data.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");

                // Encrypt the workbook with a password.
                workbook.Settings.Password = "MySecretPassword";

                // Save the encrypted workbook to a temporary file.
                tempFilePath = Path.GetTempFileName();
                workbook.Save(tempFilePath);

                // ----- Act -----
                bool exceptionThrown = false;
                try
                {
                    // Ensure the file exists before loading.
                    if (!File.Exists(tempFilePath))
                        throw new FileNotFoundException("Encrypted workbook file not found.", tempFilePath);

                    // Attempt to load the encrypted workbook without providing a password.
                    Workbook loaded = new Workbook(tempFilePath);
                }
                catch (CellsException ex) when (ex.Message.Contains("Invalid password"))
                {
                    // Expected exception for missing/incorrect password.
                    exceptionThrown = true;
                    Console.WriteLine("Invalid password exception was correctly thrown.");
                }

                // ----- Assert -----
                if (!exceptionThrown)
                {
                    Console.WriteLine("Test failed: Invalid password exception was not thrown.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions.
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
            finally
            {
                // ----- Cleanup -----
                // Delete the temporary file if it exists.
                if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                        // Suppress any errors during cleanup.
                    }
                }
            }
        }
    }
}
