// Title: Generate a CSV report that lists digitally signed and unsigned Excel workbooks in a folder using Aspose.Cells for .NET
// Common Searches: C# code to enumerate Excel files in a directory and check for digital signatures with Aspose.Cells | How to create a workbook signature compliance report in .NET using Aspose.Cells | Aspose.Cells detect signed vs unsigned .xlsx files and export results to CSV | Batch process Excel workbooks for digital signatures and generate a CSV summary in C# | Using reflection to read DigitalSignatureCollection in an Aspose.Cells workbook
// Tags: Aspose.Cells digital signature enumeration | C# CSV generation for workbook signature status | process multiple Excel extensions with Aspose.Cells | reflection based access to DigitalSignatureCollection | compliance report of signed Excel workbooks .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookSignatureReport
{
    // The sample scans a given folder for Excel workbooks (.xls, .xlsx, .xlsm, .xlsb), loads each file with Aspose.Cells, uses reflection to read the DigitalSignatureCollection count to determine if the workbook is digitally signed, separates file names into signed and unsigned lists, and writes a WorkbookSignatureReport.csv that records each file name with a Yes/No flag for signature presence, while outputting processing statistics and handling errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Specify the directory containing the workbooks
            string directoryPath = @"C:\Workbooks"; // TODO: change to target directory

            // Verify that the directory exists
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory not found: {directoryPath}");
                return;
            }

            // Prepare collections for signed and unsigned workbook file names
            List<string> signedFiles = new List<string>();
            List<string> unsignedFiles = new List<string>();

            // Define supported workbook extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Iterate through each file in the directory
            foreach (string filePath in Directory.GetFiles(directoryPath))
            {
                // Process only supported workbook files
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook using Aspose.Cells
                    Workbook workbook = new Workbook(filePath);

                    bool isSigned = false;

                    // Attempt to read DigitalSignatureCollection via reflection (avoids compile‑time dependency)
                    var dscProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                    if (dscProp != null)
                    {
                        var dsc = dscProp.GetValue(workbook);
                        var countProp = dsc?.GetType().GetProperty("Count");
                        if (countProp != null)
                        {
                            int count = (int)countProp.GetValue(dsc);
                            isSigned = count > 0;
                        }
                    }

                    // Add the file name to the appropriate list
                    if (isSigned)
                        signedFiles.Add(Path.GetFileName(filePath));
                    else
                        unsignedFiles.Add(Path.GetFileName(filePath));
                }
                catch (Exception ex)
                {
                    // If loading fails, treat as unsigned and log the error
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                    unsignedFiles.Add(Path.GetFileName(filePath));
                }
            }

            // Generate a CSV report summarizing signed vs unsigned workbooks
            string reportPath = Path.Combine(directoryPath, "WorkbookSignatureReport.csv");
            try
            {
                using (StreamWriter writer = new StreamWriter(reportPath, false))
                {
                    // Write header
                    writer.WriteLine("FileName,Signed");

                    // Write signed workbook entries
                    foreach (string fileName in signedFiles)
                    {
                        writer.WriteLine($"{fileName},Yes");
                    }

                    // Write unsigned workbook entries
                    foreach (string fileName in unsignedFiles)
                    {
                        writer.WriteLine($"{fileName},No");
                    }
                }

                Console.WriteLine($"Signature report generated at: {reportPath}");
                Console.WriteLine($"Signed workbooks: {signedFiles.Count}");
                Console.WriteLine($"Unsigned workbooks: {unsignedFiles.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }
}
