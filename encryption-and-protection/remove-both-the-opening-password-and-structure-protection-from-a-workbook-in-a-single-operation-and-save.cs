// Title: Remove opening password and workbook structure protection from an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file protected with an opening password using Aspose.Cells, clears the password, removes any workbook structure protection, and saves the file as an unprotected workbook. | Show how to combine LoadOptions and Workbook.Unprotect in Aspose.Cells to strip both file‑level and structure protection from an Excel workbook in a single operation.
// Common Searches: aspnet how to strip opening password and structure protection from Excel file using Aspose.Cells | C# remove workbook protection and opening password with Aspose.Cells LoadOptions | unprotect password protected XLSX and save unprotected copy Aspose.Cells .NET example | Aspose.Cells remove workbook structure protection programmatically | load password protected Excel workbook and save without any protection using Aspose.Cells
// Tags: Aspose.Cells remove opening password | Aspose.Cells unprotect workbook structure | Aspose.Cells load password protected XLSX | Aspose.Cells save unprotected workbook | Aspose.Cells workbook protection removal .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# example that loads a password‑protected XLSX workbook with Aspose.Cells, clears the opening password, attempts to unprotect the workbook structure, and saves the result as an unprotected Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "ProtectedWorkbook.xlsx";
            const string outputPath = "UnprotectedWorkbook.xlsx";
            const string password = "yourPassword";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook with the opening password (if known)
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };

                var workbook = new Workbook(inputPath, loadOptions);

                // Remove the opening password
                workbook.Settings.Password = string.Empty;

                // Attempt to remove workbook structure protection
                try
                {
                    // If the workbook is protected with a password, unprotect it.
                    // Passing an empty string will succeed if no password is set.
                    workbook.Unprotect(password);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to modify workbook protection settings. {ex.Message}");
                }

                // Save the unprotected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without protection: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
