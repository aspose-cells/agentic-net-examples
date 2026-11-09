// Title: Generate a signed vs unsigned Excel workbook compliance report from a folder using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells code to list Excel workbooks that have digital signatures in a folder | How to create a compliance report of signed and unsigned Excel files with Aspose.Cells | Scan a directory for signed Excel workbooks using Aspose.Cells and output a text report | Aspose.Cells digital signature detection example for .xlsx and .xlsm files in C# | Generate workbook signature summary file for multiple Excel documents with Aspose.Cells
// Tags: Aspose.Cells detect workbook digital signatures C# | C# generate signed workbook list from directory | Excel file signature compliance report using Aspose.Cells | Folder scan for signed Excel workbooks Aspose.Cells | Create text report of signed and unsigned workbooks C# | DigitalSignatureCollection reflection Aspose.Cells example

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The C# program scans a specified folder for Excel files, loads each workbook with Aspose.Cells, uses reflection to read the DigitalSignatureCollection count, classifies the files as signed or unsigned, and writes the results to a WorkbookSignatureReport.txt file for compliance review.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Determine the directory to scan – command‑line argument or default.
            string directoryPath = args.Length > 0 ? args[0] : @"C:\Workbooks";

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"Directory does not exist: {directoryPath}");
                return;
            }

            // Lists to hold signed and unsigned workbook names.
            List<string> signedWorkbooks = new List<string>();
            List<string> unsignedWorkbooks = new List<string>();

            // Supported Excel extensions.
            string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

            // Scan files in the directory.
            foreach (string filePath in Directory.GetFiles(directoryPath))
            {
                // Process only supported Excel files.
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                if (!File.Exists(filePath))
                    continue; // Safety check, though GetFiles should return existing files.

                try
                {
                    // Load workbook.
                    Workbook workbook = new Workbook(filePath);

                    // Attempt to read digital signatures via reflection (avoids compile‑time dependency on the property).
                    int signatureCount = 0;
                    var dscProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                    if (dscProp != null)
                    {
                        var dsc = dscProp.GetValue(workbook);
                        var countProp = dsc?.GetType().GetProperty("Count");
                        if (countProp != null)
                        {
                            signatureCount = (int)countProp.GetValue(dsc);
                        }
                    }

                    if (signatureCount > 0)
                        signedWorkbooks.Add(Path.GetFileName(filePath));
                    else
                        unsignedWorkbooks.Add(Path.GetFileName(filePath));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Create compliance report.
            string reportFile = Path.Combine(directoryPath, "WorkbookSignatureReport.txt");
            using (StreamWriter writer = new StreamWriter(reportFile))
            {
                writer.WriteLine("Signed Workbooks:");
                foreach (string name in signedWorkbooks)
                    writer.WriteLine(name);

                writer.WriteLine();
                writer.WriteLine("Unsigned Workbooks:");
                foreach (string name in unsignedWorkbooks)
                    writer.WriteLine(name);
            }

            Console.WriteLine($"Compliance report generated: {reportFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
