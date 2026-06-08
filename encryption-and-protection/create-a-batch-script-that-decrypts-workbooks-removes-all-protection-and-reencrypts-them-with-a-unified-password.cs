using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookBatchProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to process
            string sourceFolder = @"C:\InputWorkbooks";
            // Folder where the processed workbooks will be saved
            string destinationFolder = @"C:\OutputWorkbooks";

            // Original password used to open the encrypted workbooks (if any)
            string originalPassword = "oldPassword";
            // Unified password to re‑encrypt all workbooks
            string unifiedPassword = "newUnifiedPassword";

            // Ensure the destination folder exists
            Directory.CreateDirectory(destinationFolder);

            // Process each workbook file in the source folder
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                // Consider only Excel related extensions
                string extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm" && extension != ".xlsb" && extension != ".ods")
                    continue;

                try
                {
                    // Load the workbook (use password if the file is encrypted)
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = originalPassword; // If the workbook is not encrypted this will be ignored
                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // ----- Remove workbook protection -----
                    if (workbook.IsWorkbookProtectedWithPassword)
                    {
                        workbook.Unprotect(originalPassword);
                    }

                    // ----- Remove shared workbook protection (if any) -----
                    // Attempt to unprotect; if not protected an exception will not be thrown
                    try
                    {
                        workbook.UnprotectSharedWorkbook(originalPassword);
                    }
                    catch { /* ignore if not a shared workbook */ }

                    // ----- Remove worksheet protection for each sheet -----
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.IsProtected)
                        {
                            sheet.Unprotect(originalPassword);
                        }
                    }

                    // ----- Optional cleanup (remove macros, digital signatures, personal info) -----
                    // Uncomment if required
                    // workbook.RemoveMacro();
                    // workbook.RemoveDigitalSignature();
                    // workbook.RemovePersonalInformation();

                    // ----- Re‑encrypt the workbook with the unified password -----
                    workbook.Settings.Password = unifiedPassword;

                    // Save the processed workbook to the destination folder (overwrite if exists)
                    string destPath = Path.Combine(destinationFolder, Path.GetFileName(filePath));
                    workbook.Save(destPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}