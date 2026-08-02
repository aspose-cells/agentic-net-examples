// Title: Retry Decryption with Alternate Password Using Aspose.Cells for .NET
// Description: Shows how to open an encrypted Excel workbook with a primary password, catch the failure, retry using a secondary password, then strip protection and save the file unencrypted, while handling missing files and other errors.
// Keywords: Aspose.Cells password retry | C# load encrypted workbook fallback | Aspose.Cells decryption error handling | remove Excel password Aspose.Cells | .NET workbook encryption retry | LoadOptions password alternative
// Common Searches: Aspose.Cells retry opening encrypted Excel with another password | C# load workbook with fallback password after failure | How to remove password protection after loading encrypted file Aspose.Cells | FileNotFoundException handling before decrypting Excel in C# | Aspose.Cells error handling for wrong password
// Developer Intent: Implement robust error handling that attempts to open an encrypted workbook with a primary password and automatically retries with an alternate password if the first attempt fails.
// Use Cases: Attempt decryption when the exact password is uncertain by trying multiple candidates sequentially. | Automatically remove workbook protection after a successful load and save a clean copy. | Validate the existence of the target Excel file before any decryption attempt to avoid runtime exceptions.
// AI Prompts: Generate C# code that accepts a file path and an array of passwords, loads the workbook with the first valid password using Aspose.Cells, and returns the workbook object. | Refactor the retry logic into a reusable method named LoadWorkbookWithFallback that logs each attempt and throws a custom exception if all passwords fail. | Create unit tests for LoadWorkbookWithFallback that mock successful decryption with the second password and verify proper exception handling for missing files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to open an encrypted Excel workbook with a primary password, catch the failure, retry using a secondary password, then strip protection and save the file unencrypted, while handling missing files and other errors.
    public class DecryptionRetryDemo
    {
        public static void Run()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Verify that the file exists to prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file '{filePath}' was not found.");
                return;
            }

            // First (incorrect) password attempt
            string primaryPassword = "wrongPassword";

            // Alternate password to try if the first attempt fails
            string alternatePassword = "correctPassword";

            Workbook workbook = null;

            // Attempt to load the workbook with the primary password
            try
            {
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = primaryPassword
                };

                workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully with primary password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Primary password failed: {ex.Message}");
                Console.WriteLine("Retrying with alternate password...");

                // Retry with the alternate password
                try
                {
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = alternatePassword
                    };

                    workbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully with alternate password.");
                }
                catch (Exception retryEx)
                {
                    Console.WriteLine($"Alternate password also failed: {retryEx.Message}");
                    // If both attempts fail, exit the method
                    return;
                }
            }

            // At this point, 'workbook' is loaded successfully.
            // Example operation: remove the password protection and save unencrypted.
            try
            {
                workbook.Settings.Password = null; // Remove encryption password
                workbook.Save("unprotected.xlsx");
                Console.WriteLine("Workbook saved without password protection.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                DecryptionRetryDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
