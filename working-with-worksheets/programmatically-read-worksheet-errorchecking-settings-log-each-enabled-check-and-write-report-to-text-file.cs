// Title: C# – Read Worksheet Error‑Checking Settings with Aspose.Cells and Export to Text Report
// Description: Load an Excel workbook with Aspose.Cells for .NET, access each worksheet's ErrorCheckOptionCollection, list enabled ErrorCheckType values (including range details), and write a concise text file summarizing the error‑checking configuration.
// Keywords: Aspose.Cells read error check options | C# Excel error checking report | ErrorCheckOptionCollection example | export worksheet error checks to txt | list enabled ErrorCheckType Aspose
// Common Searches: how to get error‑check settings from a worksheet using Aspose.Cells | export Excel error checking configuration to a text file C# | iterate ErrorCheckOptionCollection Aspose.Cells .NET | retrieve ranges for error‑check options in a workbook
// Developer Intent: Extract enabled error‑checking rules from a worksheet and save them as a readable text report.
// Use Cases: Audit workbooks for data‑validation problems by listing active error checks per sheet. | Create compliance documentation that records each worksheet's error‑checking configuration. | Build a diagnostic tool that scans multiple Excel files and logs their error‑check settings for quality control.
// AI Prompts: Generate a method that accepts a Workbook and returns a formatted error‑check report for all its worksheets using Aspose.Cells. | Extend the sample to loop through every worksheet in the workbook and append each sheet’s report to a single text file. | Add comprehensive error handling to manage missing files, empty worksheets, or unsupported error‑check types while creating the report.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// Load an Excel workbook with Aspose.Cells for .NET, access each worksheet's ErrorCheckOptionCollection, list enabled ErrorCheckType values (including range details), and write a concise text file summarizing the error‑checking configuration.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can loop through all worksheets if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare a list to hold report lines
        List<string> report = new List<string>();
        report.Add($"Worksheet: {worksheet.Name}");
        report.Add("Enabled Error Checks:");

        // Get the collection of error‑check options for the worksheet
        ErrorCheckOptionCollection options = worksheet.ErrorCheckOptions;

        // Iterate through each ErrorCheckOption in the collection
        for (int i = 0; i < options.Count; i++)
        {
            ErrorCheckOption option = options[i];

            // Determine the ranges this option applies to
            int rangeCount = option.GetCountOfRange();
            string rangeInfo = rangeCount > 0 ? $"Ranges ({rangeCount})" : "No specific range";

            // Collect all enabled error‑check types for this option
            List<string> enabledTypes = new List<string>();
            foreach (ErrorCheckType type in Enum.GetValues(typeof(ErrorCheckType)))
            {
                if (option.IsErrorCheck(type))
                {
                    enabledTypes.Add(type.ToString());
                }
            }

            // If any checks are enabled, add them to the report
            if (enabledTypes.Count > 0)
            {
                report.Add($"Option {i}: {rangeInfo}");
                foreach (string typeName in enabledTypes)
                {
                    report.Add($"  - {typeName}");
                }
            }
        }

        // Write the report to a text file
        string reportPath = "ErrorCheckReport.txt";
        File.WriteAllLines(reportPath, report);
        Console.WriteLine($"Error‑check report written to {reportPath}");
    }
}
