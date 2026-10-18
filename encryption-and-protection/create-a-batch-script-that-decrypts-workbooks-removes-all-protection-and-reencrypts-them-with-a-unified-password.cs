// Title: C# batch script to decrypt, remove protection, and re‑encrypt multiple Excel .xlsx workbooks with a single password using Aspose.Cells
// Common Searches: C# batch program to remove workbook and worksheet protection from many .xlsx files with Aspose.Cells | How to load encrypted Excel files, unprotect them, and save with a new password using Aspose.Cells in .NET | Aspose.Cells script to iterate over a folder of Excel workbooks, clear protection, and apply a common password | Re‑encrypt multiple Excel workbooks with a unified password after removing existing protection in C# | Automate decryption and re‑encryption of protected Excel files using Aspose.Cells LoadOptions password
// Tags: decrypt Excel workbook Aspose.Cells C# | remove workbook protection Aspose.Cells | unprotect worksheet Aspose.Cells programmatically | apply unified password to .xlsx files Aspose.Cells | batch process multiple Excel files Aspose.Cells | load encrypted workbook with password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookProtectionBatch
{
    // A C# console utility that accepts input and output folders plus optional old and new passwords, scans all .xlsx files, loads each workbook with the old password, removes workbook‑level and worksheet‑level protection, sets a new unified password, and saves the re‑encrypted files to the output directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Expecting arguments: inputFolder outputFolder [oldPassword] [newPassword]
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: WorkbookProtectionBatch <inputFolder> <outputFolder> [oldPassword] [newPassword]");
                return;
            }

            string inputFolder = args[0];
            string outputFolder = args[1];
            string oldPassword = args.Length > 2 ? args[2] : string.Empty;
            string newPassword = args.Length > 3 ? args[3] : "UnifiedPassword";

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Process each Excel file in the input folder
            string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories);
            foreach (string filePath in files)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load workbook with the old password (if any)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    if (!string.IsNullOrEmpty(oldPassword))
                        loadOptions.Password = oldPassword;

                    Workbook workbook = new Workbook(filePath, loadOptions);

                    // Remove workbook-level protection (structure, windows) if present
                    try
                    {
                        workbook.Unprotect(oldPassword);
                    }
                    catch (Exception ex)
                    {
                        // Ignore if workbook was not protected or password is incorrect
                        Console.WriteLine($"Warning: Could not unprotect workbook '{Path.GetFileName(filePath)}' - {ex.Message}");
                    }

                    // Remove protection from each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        try
                        {
                            sheet.Unprotect(oldPassword);
                        }
                        catch (Exception ex)
                        {
                            // Ignore if sheet was not protected or password is incorrect
                            Console.WriteLine($"Warning: Could not unprotect sheet '{sheet.Name}' in '{Path.GetFileName(filePath)}' - {ex.Message}");
                        }
                    }

                    // Set the new unified password for encryption
                    workbook.Settings.Password = newPassword;

                    // Save the workbook to the output folder (encrypted with new password)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process {Path.GetFileName(filePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
