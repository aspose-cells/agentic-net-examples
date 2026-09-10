// Title: Generate a signed vs unsigned Excel workbook report from a folder using Aspose.Cells in C#
// AI Prompts: Write a C# console program that enumerates .xls, .xlsx, .xlsm, and .xlsb files in a given path, loads each workbook with Aspose.Cells, inspects the DigitalSignatureCollection via reflection, and writes two sections (Signed and Unsigned) to a plain‑text report. | Extend the workbook scanner to walk subfolders recursively and output the signature results as a CSV file with columns FileName and SignatureStatus. | Add comprehensive error handling that records files which cannot be opened to a separate log and appends a summary line showing total counts of signed and unsigned workbooks.
// Common Searches: c# enumerate Excel files and detect digital signatures with Aspose.Cells | how to create a compliance report of signed versus unsigned workbooks using Aspose.Cells .NET | scan a folder for .xls/.xlsx files and list those with digital signatures in C# | generate text report of signed and unsigned Excel workbooks programmatically | Aspose.Cells check workbook DigitalSignatureCollection in a .NET console app
// Tags: Aspose.Cells digital signature detection .NET | C# enumerate Excel workbooks by signature status | generate signed workbook report Aspose.Cells | scan folder for signed .xlsx files using Aspose.Cells | write workbook signature summary to text file C#

using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorkbookSignatureReport
{
    // // This C# console application scans a specified directory for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, uses reflection to examine the DigitalSignatureCollection for signatures, separates file names into signed and unsigned lists, and writes the results to WorkbookSignatureReport.txt.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the workbooks
            string directoryPath = @"C:\Workbooks";

            // Verify that the directory exists
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory not found: {directoryPath}");
                return;
            }

            // Supported workbook extensions
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Lists to store signed and unsigned workbook names
            List<string> signedWorkbooks = new List<string>();
            List<string> unsignedWorkbooks = new List<string>();

            try
            {
                // Process each extension separately
                foreach (string ext in extensions)
                {
                    // Get files with the current extension
                    string[] files = Directory.GetFiles(directoryPath, "*" + ext, SearchOption.TopDirectoryOnly);

                    foreach (string filePath in files)
                    {
                        // Ensure the file exists before attempting to load
                        if (!File.Exists(filePath))
                            continue;

                        try
                        {
                            // Load workbook
                            Workbook workbook = new Workbook(filePath);

                            // Determine if the workbook has digital signatures using reflection
                            bool hasSignature = false;
                            var propInfo = workbook.GetType().GetProperty("DigitalSignatureCollection");
                            if (propInfo != null)
                            {
                                var collection = propInfo.GetValue(workbook) as ICollection;
                                if (collection != null && collection.Count > 0)
                                    hasSignature = true;
                            }

                            // Add file name to the appropriate list
                            if (hasSignature)
                                signedWorkbooks.Add(Path.GetFileName(filePath));
                            else
                                unsignedWorkbooks.Add(Path.GetFileName(filePath));
                        }
                        catch (Exception ex)
                        {
                            // Log loading errors and continue processing other files
                            Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                        }
                    }
                }

                // Write the report
                string reportPath = Path.Combine(directoryPath, "WorkbookSignatureReport.txt");
                using (StreamWriter writer = new StreamWriter(reportPath))
                {
                    writer.WriteLine("Signed Workbooks:");
                    foreach (string name in signedWorkbooks)
                        writer.WriteLine(name);

                    writer.WriteLine(); // Separator

                    writer.WriteLine("Unsigned Workbooks:");
                    foreach (string name in unsignedWorkbooks)
                        writer.WriteLine(name);
                }

                Console.WriteLine("Workbook signature report generated at:");
                Console.WriteLine(reportPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
