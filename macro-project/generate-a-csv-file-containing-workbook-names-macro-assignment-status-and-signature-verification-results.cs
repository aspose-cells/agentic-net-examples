// Title: Generate a CSV report of macro presence and digital signature verification for Excel workbooks using Aspose.Cells in C#
// AI Prompts: Write a C# console application that scans a folder for .xls, .xlsx, .xlsm, and .xlsb files, loads each workbook with Aspose.Cells, determines whether VBA macros exist, checks digital signatures via the DigitalSignatureCollection (using reflection for version‑agnostic support), and writes the workbook name, macro flag, and signature status to a UTF‑8 CSV file. | Update the program to gracefully handle environments where the DigitalSignatureCollection API is unavailable, output "Not Supported" for signature status in those cases, and ensure the CSV header columns are Workbook Name, Has Macros, and Signature Verification.
// Common Searches: c# Aspose.Cells batch export macro detection results to CSV | how to check VBA macros in multiple Excel files with Aspose.Cells .NET | retrieve digital signature status of Excel workbooks using Aspose.Cells reflection | generate UTF-8 CSV report of workbook name macro presence signature verification Aspose.Cells | process .xlsb files for macro and signature info with Aspose.Cells
// Tags: export macro detection results to CSV using Aspose.Cells | detect VBA modules in Excel workbooks .NET | verify Excel digital signatures via Aspose.Cells reflection | batch process .xlsb .xlsx files for macro status Aspose.Cells | Aspose.Cells workbook metadata extraction to CSV

using System;
using System.IO;
using System.Text;
using System.Collections;
using Aspose.Cells;

// A C# console utility that iterates over a directory of Excel files (.xls, .xlsx, .xlsm, .xlsb), loads each workbook with Aspose.Cells, identifies the presence of VBA macros, evaluates digital signature validity (using reflection to stay compatible with older library versions), and writes a UTF‑8 CSV containing the workbook name, macro presence (Yes/No), and signature verification result (Valid, Invalid, No Signatures, Not Supported).
class WorkbookInfoExporter
{
    static void Main()
    {
        // Folder containing the Excel workbooks to process
        string inputFolder = @"C:\Workbooks";

        // Path for the generated CSV file
        string outputCsvPath = @"C:\WorkbookReport.csv";

        // Ensure the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Prepare a StringBuilder for CSV content
        StringBuilder csvBuilder = new StringBuilder();

        // Write CSV header
        csvBuilder.AppendLine("Workbook Name,Has Macros,Signature Verification");

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Consider only Excel file extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                continue;

            // Verify the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Determine macro assignment status
                bool hasMacros = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;
                string macroStatus = hasMacros ? "Yes" : "No";

                // Determine signature verification result (using reflection to stay compatible with versions lacking the API)
                string signatureStatus = "No Signatures";
                try
                {
                    var dsProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                    var dsCollection = dsProp?.GetValue(workbook, null);
                    if (dsCollection != null)
                    {
                        var countProp = dsCollection.GetType().GetProperty("Count");
                        int count = (int)countProp.GetValue(dsCollection);
                        if (count > 0)
                        {
                            bool allValid = true;
                            foreach (object signature in (IEnumerable)dsCollection)
                            {
                                var verifyMethod = signature.GetType().GetMethod("Verify");
                                if (verifyMethod != null)
                                {
                                    bool isValid = (bool)verifyMethod.Invoke(signature, null);
                                    if (!isValid)
                                    {
                                        allValid = false;
                                        break;
                                    }
                                }
                            }
                            signatureStatus = allValid ? "Valid" : "Invalid";
                        }
                    }
                    else
                    {
                        signatureStatus = "Not Supported";
                    }
                }
                catch
                {
                    // If any reflection step fails, mark as not supported
                    signatureStatus = "Not Supported";
                }

                // Append a line for the current workbook
                string workbookName = Path.GetFileName(filePath);
                csvBuilder.AppendLine($"{workbookName},{macroStatus},{signatureStatus}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        try
        {
            // Save the CSV content to file
            File.WriteAllText(outputCsvPath, csvBuilder.ToString(), Encoding.UTF8);
            Console.WriteLine($"Report generated at: {outputCsvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write CSV file: {ex.Message}");
        }
    }
}
