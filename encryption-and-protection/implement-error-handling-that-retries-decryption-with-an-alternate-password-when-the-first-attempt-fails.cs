// Title: Retry opening an encrypted Excel workbook with a fallback password using Aspose.Cells for .NET
// Description: Demonstrates how to verify a protected Excel file, attempt to load it with a primary password, catch the failure, switch to an alternate password, reload the workbook, remove its protection, and save an unprotected copy. Includes comprehensive error handling for each step.
// Keywords: Aspose.Cells password retry | C# load encrypted workbook | fallback password Aspose.Cells | Excel decryption exception handling | remove workbook protection .NET | LoadOptions alternate password | retry logic encrypted Excel | Aspose.Cells error handling
// Common Searches: how to retry opening a password protected Excel file with Aspose.Cells | Aspose.Cells load workbook with secondary password | C# remove password from encrypted Excel using Aspose.Cells | exception handling for wrong password Aspose.Cells | fallback password for encrypted workbook .NET
// Developer Intent: Open a password‑protected Excel file, automatically retry with a second password if the first fails, then strip the protection and save the workbook without a password.
// Use Cases: Batch processing of multiple encrypted workbooks where the correct password may vary. | Automated migration of secured Excel files to unprotected versions for downstream systems. | Implementing resilient password handling in a data‑import pipeline that must continue despite incorrect credentials.
// AI Prompts: Write C# code using Aspose.Cells that tries a primary password, falls back to a secondary password on failure, and saves the workbook without protection. | Create a reusable Aspose.Cells method that accepts a file path, primary and secondary passwords, and returns an unprotected Workbook with proper exception handling. | Explain how to log detailed error information when both password attempts fail while loading an encrypted workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to verify a protected Excel file, attempt to load it with a primary password, catch the failure, switch to an alternate password, reload the workbook, remove its protection, and save an unprotected copy. Includes comprehensive error handling for each step.
    public class DecryptionRetryDemo
    {
        public static void Run()
        {
            // Path to the encrypted workbook
            string filePath = "protected.xlsx";

            // Verify the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Input file not found: {filePath}");
                return;
            }

            // First password attempt (may be incorrect)
            string primaryPassword = "wrongPassword";

            // Alternate password to try if the first one fails
            string alternatePassword = "correctPassword";

            // LoadOptions will hold the password for opening the workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = primaryPassword
            };

            Workbook workbook = null;

            try
            {
                // Attempt to load the workbook with the primary password
                workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully with primary password.");
            }
            catch (Exception exPrimary)
            {
                // Loading failed – likely due to an incorrect password
                Console.WriteLine($"Primary password failed: {exPrimary.Message}");
                Console.WriteLine("Retrying with alternate password...");

                try
                {
                    // Set the alternate password and retry loading
                    loadOptions.Password = alternatePassword;
                    workbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully with alternate password.");
                }
                catch (Exception exAlternate)
                {
                    Console.WriteLine($"Alternate password failed: {exAlternate.Message}");
                    return;
                }
            }

            // Ensure workbook was loaded before proceeding
            if (workbook == null)
            {
                Console.WriteLine("Failed to load workbook.");
                return;
            }

            // Remove the password protection after successful load
            workbook.Settings.Password = null;

            // Save the unprotected workbook
            string outputPath = "unprotected.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Unprotected workbook saved to '{outputPath}'.");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
    }

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
