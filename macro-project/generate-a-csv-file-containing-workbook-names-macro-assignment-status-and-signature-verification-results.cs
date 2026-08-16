// Title: C# – Generate CSV Report of Excel Macro Presence and Digital/VBA Signature Status with Aspose.Cells
// Description: A concise C# utility that scans a list of Excel files, uses Aspose.Cells to detect macros, checks if the workbook and its VBA project are digitally signed, validates the VBA signature, and writes the results (Workbook, HasMacro, IsDigitallySigned, VbaSigned, ValidVbaSignature) to a CSV file.
// Keywords: Aspose.Cells C# macro detection | Excel digital signature verification .NET | VBA project signature validation | generate CSV report Aspose.Cells | audit Excel workbooks for macros | C# Excel security report
// Common Searches: Aspose.Cells list workbooks with macros and signatures | C# create CSV of Excel macro and digital signature status | check VBA signature validity using Aspose.Cells | how to detect signed Excel files in .NET | generate macro audit report for Excel files
// Developer Intent: Produce a CSV file that enumerates each workbook’s name together with flags indicating macro presence, digital signing, VBA signing, and VBA signature validity.
// Use Cases: Compliance audit of Excel assets to enforce macro and signing policies. | Security reporting to identify unsigned or tampered VBA projects. | Automated governance for IT departments tracking macro usage across multiple workbooks.
// AI Prompts: Write C# code with Aspose.Cells that reads a collection of Excel files and outputs a CSV containing Workbook, HasMacro, IsDigitallySigned, VbaSigned, ValidVbaSignature. | Extend the sample to add the workbook's creation date as an extra column in the CSV report. | Add robust logging that records missing files and signature verification errors to a separate log while continuing processing.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// A concise C# utility that scans a list of Excel files, uses Aspose.Cells to detect macros, checks if the workbook and its VBA project are digitally signed, validates the VBA signature, and writes the results (Workbook, HasMacro, IsDigitallySigned, VbaSigned, ValidVbaSignature) to a CSV file.
class WorkbookReportGenerator
{
    static void Main()
    {
        // List of workbook file paths to process
        string[] workbookPaths = new string[]
        {
            "Sample1.xlsx",
            "Sample2.xlsm",
            "Sample3.xlsx"
        };

        // Prepare CSV header
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Workbook,HasMacro,IsDigitallySigned,VbaSigned,ValidVbaSignature");

        foreach (string path in workbookPaths)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}. Skipping.");
                continue;
            }

            try
            {
                // Load the workbook inside a using block to ensure proper disposal
                using (Workbook wb = new Workbook(path))
                {
                    // Extract workbook name
                    string workbookName = Path.GetFileName(path);

                    // Macro assignment status
                    bool hasMacro = wb.HasMacro;

                    // Digital signature status
                    bool isDigitallySigned = wb.IsDigitallySigned;

                    // VBA project signature status (if VBA project exists)
                    bool vbaSigned = wb.VbaProject != null && wb.VbaProject.IsSigned;
                    bool vbaValidSigned = vbaSigned && wb.VbaProject.IsValidSigned;

                    // Append a CSV line with the collected data
                    csvBuilder.AppendLine($"{workbookName},{hasMacro},{isDigitallySigned},{vbaSigned},{vbaValidSigned}");
                }
            }
            catch (Exception ex)
            {
                // Log the error and continue with the next file
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }

        try
        {
            // Write the CSV content to a file
            File.WriteAllText("WorkbookReport.csv", csvBuilder.ToString());
            Console.WriteLine("CSV report generated: WorkbookReport.csv");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write CSV report: {ex.Message}");
        }
    }
}
