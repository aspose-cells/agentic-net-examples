// Title: Create a CSV report of Excel workbook names, macro detection, and signature verification status using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to export Excel file names and macro detection to CSV | How to list workbooks in a folder and check for VBA macros using Aspose.Cells | Generate workbook summary CSV with macro presence and digital signature check in .NET | Aspose.Cells .NET example for iterating .xls files and writing macro status to CSV | Export Excel macro detection results to CSV using Aspose.Cells library
// Tags: Aspose.Cells VBA project presence extraction C# | CSV generation for Excel workbook inventory Aspose.Cells | identify VBA modules in .xls files using Aspose.Cells | signature verification not supported flag Aspose.Cells | batch workbook processing to summary CSV .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a directory for Excel files, loads each workbook with Aspose.Cells, determines if a VBA project with modules exists, marks digital signature verification as NotSupported, and writes the workbook name, macro presence, and signature result into a CSV report.
class WorkbookInfoExporter
{
    static void Main(string[] args)
    {
        // Folder containing the workbooks to process
        string sourceFolder = @"C:\Workbooks";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Output CSV file path
        string csvPath = @"C:\WorkbookReport.csv";

        try
        {
            // Prepare the CSV file with header
            using (var writer = new StreamWriter(csvPath, false))
            {
                writer.WriteLine("WorkbookName,HasMacros,SignatureVerification");

                // Enumerate all Excel files in the source folder
                foreach (string filePath in Directory.GetFiles(sourceFolder, "*.xls*"))
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook using Aspose.Cells
                        Workbook workbook = new Workbook(filePath);

                        // Determine macro assignment status
                        bool hasMacros = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;

                        // Digital signature verification is not available in the current Aspose.Cells version.
                        // Set result to "NotSupported" to indicate this.
                        string signatureResult = "NotSupported";

                        // Write the information to the CSV
                        string workbookName = Path.GetFileName(filePath);
                        writer.WriteLine($"{workbookName},{(hasMacros ? "Yes" : "No")},{signatureResult}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"Workbook report generated at: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to generate report: {ex.Message}");
        }
    }
}
