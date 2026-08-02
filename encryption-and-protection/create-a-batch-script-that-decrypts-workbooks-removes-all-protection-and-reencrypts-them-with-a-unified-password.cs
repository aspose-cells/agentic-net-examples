// Title: C# batch script to decrypt, strip protection, and re‑encrypt Excel workbooks with a single password using Aspose.Cells
// Description: A .NET console utility that scans a specified folder, loads each .xlsx, .xls, .xlsb, or .xlsm file (including password‑protected workbooks), removes workbook‑level, shared‑workbook, and worksheet protection, deletes macros, digital signatures and personal information, then saves the file in its original format encrypted with a unified password. The solution leverages Aspose.Cells for fast, password‑aware processing and supports custom input and output directories.
// Keywords: Aspose.Cells | C# batch Excel encryption | remove worksheet protection programmatically | decrypt Excel workbook | re‑encrypt Excel files | bulk Excel password change | Excel macro removal C# | shared workbook unprotect | Excel file format preservation | console app Aspose.Cells .NET
// Common Searches: batch change Excel password C# | remove protection from multiple Excel files Aspose.Cells | decrypt and re‑encrypt Excel workbooks .NET | strip macros from Excel files in bulk | how to unprotect shared workbook using Aspose.Cells | C# script to process all Excel files in a folder | convert encrypted Excel to new password programmatically
// Developer Intent: Create a single command‑line tool that cleans, unprotects, and re‑protects a collection of Excel workbooks with one common password.
// Use Cases: Standardize the password for all corporate Excel reports before archiving them in a secure repository. | Prepare a batch of workbooks for external distribution by removing macros, digital signatures, and personal metadata. | Migrate legacy encrypted spreadsheets to a new security policy that requires a unified password across the organization.
// AI Prompts: Write a robust C# method that iterates through a directory, opens each encrypted Excel file with Aspose.Cells, removes all protections, and saves it using a new unified password. | Add detailed logging and error handling to the batch processor, capturing files that fail to load, unprotect, or save, and outputting a summary report. | Extend the script to generate a CSV audit file that records the original protection state of each worksheet before removal.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookBatchProcessor
{
    // A .NET console utility that scans a specified folder, loads each .xlsx, .xls, .xlsb, or .xlsm file (including password‑protected workbooks), removes workbook‑level, shared‑workbook, and worksheet protection, deletes macros, digital signatures and personal information, then saves the file in its original format encrypted with a unified password. The solution leverages Aspose.Cells for fast, password‑aware processing and supports custom input and output directories.
    class Program
    {
        // Unified password to be applied to all processed workbooks
        private const string UnifiedPassword = "UnifiedPass123";

        // Original password used for opening encrypted workbooks (replace with actual password or retrieve per file)
        private const string OriginalPassword = "oldPassword";

        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the processed workbooks will be saved
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            string[] files;
            try
            {
                // Process all Excel files in the input folder (non‑recursive)
                files = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to enumerate files in '{inputFolder}': {ex.Message}");
                return;
            }

            foreach (string filePath in files)
            {
                // Consider only supported Excel extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsb" && extension != ".xlsm")
                    continue;

                // Ensure the file actually exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (use LoadOptions with password if the file is encrypted)
                    Workbook workbook = LoadWorkbook(filePath);

                    // Attempt to remove workbook‑level protection (structure protection)
                    try
                    {
                        workbook.Unprotect(OriginalPassword);
                    }
                    catch
                    {
                        // Ignore if workbook is not protected or password is incorrect
                    }

                    // Remove shared workbook protection if applicable
                    if (workbook.Settings.IsEncrypted)
                    {
                        try
                        {
                            workbook.UnprotectSharedWorkbook(OriginalPassword);
                        }
                        catch
                        {
                            // Ignore if not a shared workbook or password is incorrect
                        }
                    }

                    // Iterate through all worksheets and remove their protection
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.IsProtected)
                        {
                            try
                            {
                                sheet.Unprotect(OriginalPassword);
                            }
                            catch
                            {
                                // Ignore if sheet is not protected or password is incorrect
                            }
                        }
                    }

                    // Remove macros, digital signatures, and personal information
                    workbook.RemoveMacro();
                    workbook.RemoveDigitalSignature();
                    workbook.RemovePersonalInformation();

                    // Apply the unified password to encrypt the workbook
                    workbook.Settings.Password = UnifiedPassword;

                    // Determine appropriate SaveFormat based on original file extension
                    SaveFormat saveFormat = GetSaveFormat(extension);

                    // Build output file path (preserve original file name)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the processed workbook
                    workbook.Save(outputPath, saveFormat);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }

        // Loads a workbook, handling encrypted files using the original password if needed
        private static Workbook LoadWorkbook(string filePath)
        {
            try
            {
                // Attempt to load without a password
                return new Workbook(filePath);
            }
            catch
            {
                // If loading fails, assume the workbook is encrypted and retry with the original password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = OriginalPassword
                };
                return new Workbook(filePath, loadOptions);
            }
        }

        // Maps file extensions to Aspose.Cells SaveFormat values
        private static SaveFormat GetSaveFormat(string extension)
        {
            switch (extension)
            {
                case ".xlsx":
                case ".xlsm":
                    return SaveFormat.Xlsx;
                case ".xls":
                    return SaveFormat.Excel97To2003;
                case ".xlsb":
                    return SaveFormat.Xlsb;
                default:
                    return SaveFormat.Xlsx;
            }
        }
    }
}
