// Title: Detect Default Excel Encryption with Aspose.Cells .NET and Recommend Upgrade
// Description: C# sample that loads an Excel workbook using Aspose.Cells, checks Workbook.Settings.IsEncrypted and IsDefaultEncrypted, reports the encryption status, and suggests applying a stronger custom encryption algorithm when the default is detected.
// Keywords: Aspose.Cells default encryption detection | Workbook.Settings.IsDefaultEncrypted | C# Excel encryption check | upgrade Excel encryption Aspose | log encryption recommendation .NET
// Common Searches: how to know if Excel file uses Aspose.Cells default encryption | Aspose.Cells check if workbook is encrypted with default algorithm | C# code to recommend stronger encryption for Excel files | detect default encryption in .xlsx using Aspose.Cells
// Developer Intent: Determine whether an Excel workbook is protected with Aspose.Cells' default encryption and output a recommendation to switch to a custom, stronger encryption method.
// Use Cases: Run a security audit across a repository of Excel files to flag workbooks that rely on the default encryption. | Integrate the check into a CI/CD pipeline to fail builds when default encryption is found. | Add a runtime warning in an application that opens workbooks, prompting users to upgrade the encryption settings.
// AI Prompts: Write C# code that scans a directory for .xlsx files, uses Aspose.Cells to detect default encryption, and creates a CSV report with upgrade suggestions. | Refactor the EncryptionDetector class to expose a reusable method returning a boolean for default encryption and optionally apply custom encryption options. | Show how to set a custom password and encryption level with Aspose.Cells SetEncryptionOptions after detecting default encryption.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionCheck
{
    // C# sample that loads an Excel workbook using Aspose.Cells, checks Workbook.Settings.IsEncrypted and IsDefaultEncrypted, reports the encryption status, and suggests applying a stronger custom encryption algorithm when the default is detected.
    public class EncryptionDetector
    {
        // Detects if the workbook uses the default encryption algorithm and logs a recommendation.
        public static void Run(string workbookPath)
        {
            // Prevent FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook inside a using block to ensure disposal
                using (Workbook workbook = new Workbook(workbookPath))
                {
                    bool isEncrypted = workbook.Settings.IsEncrypted;
                    bool isDefaultEncrypted = workbook.Settings.IsDefaultEncrypted;

                    if (isEncrypted && isDefaultEncrypted)
                    {
                        Console.WriteLine($"Workbook '{workbookPath}' is encrypted using the default encryption algorithm.");
                        Console.WriteLine("Recommendation: Upgrade to a stronger encryption algorithm (e.g., set a custom password and use SetEncryptionOptions).");
                    }
                    else if (isEncrypted)
                    {
                        Console.WriteLine($"Workbook '{workbookPath}' is encrypted with a custom algorithm.");
                    }
                    else
                    {
                        Console.WriteLine($"Workbook '{workbookPath}' is not encrypted.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        // Entry point required for compilation
        static void Main(string[] args)
        {
            string workbookPath;

            if (args.Length > 0)
            {
                workbookPath = args[0];
            }
            else
            {
                Console.Write("Enter the path to the Excel workbook: ");
                workbookPath = Console.ReadLine();
            }

            EncryptionDetector.Run(workbookPath);
        }
    }
}
