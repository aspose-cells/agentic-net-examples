// Title: C# – Create CSV Report of Workbook CheckCompatibility Flag and Ribbon XML with Aspose.Cells
// Description: Loads a set of Excel files, reads each workbook's Settings.CheckCompatibility property and RibbonXml presence, and writes a CSV file that lists the file name, compatibility status, and whether custom Ribbon XML is defined.
// Keywords: Aspose.Cells C# read CheckCompatibility | Aspose.Cells RibbonXml detection | batch workbook analysis Aspose.Cells | export Excel settings to CSV | Aspose.Cells workbook audit
// Common Searches: how to get CheckCompatibility using Aspose.Cells .NET | list workbooks with RibbonXml via Aspose.Cells | generate CSV of Excel workbook settings Aspose.Cells | Aspose.Cells batch process multiple workbooks
// Developer Intent: Generate a CSV summary that shows each workbook’s filename, its compatibility flag, and whether custom Ribbon XML is set.
// Use Cases: Verify compatibility settings across a large collection of Excel files before release. | Audit custom UI customizations by identifying workbooks that contain Ribbon XML. | Create an inventory of workbook properties for migration, compliance, or quality‑control purposes.
// AI Prompts: Write C# code with Aspose.Cells that reads Settings.CheckCompatibility and RibbonXml for multiple Excel files and outputs the results to a CSV file. | Show how to handle missing files gracefully while continuing to process the remaining workbooks in Aspose.Cells. | Suggest extensions to the CSV report to include additional properties such as EnableMacros, IsProtected, or WorkbookVersion.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Loads a set of Excel files, reads each workbook's Settings.CheckCompatibility property and RibbonXml presence, and writes a CSV file that lists the file name, compatibility status, and whether custom Ribbon XML is defined.
class Program
{
    static void Main()
    {
        // Define the workbook files to process
        var workbookPaths = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsm"
            // Add more file paths as needed
        };

        // Prepare a list to hold report lines (CSV format)
        var reportLines = new List<string>();
        reportLines.Add("Workbook,CheckCompatibility,RibbonXmlSet");

        foreach (var path in workbookPaths)
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(path))
            {
                Console.WriteLine($"File not found: {path}");
                continue;
            }

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(path);

            // Retrieve the compatibility setting from WorkbookSettings
            bool checkCompatibility = workbook.Settings.CheckCompatibility;

            // Determine whether RibbonXml has been set (non‑null and non‑empty)
            bool ribbonXmlSet = !string.IsNullOrEmpty(workbook.RibbonXml);

            // Add a line to the report
            reportLines.Add($"{Path.GetFileName(path)},{checkCompatibility},{ribbonXmlSet}");

            // Release resources
            workbook.Dispose();
        }

        // Write the summary report to a CSV file (uses the provided save rule)
        string reportPath = "WorkbookSummaryReport.csv";
        File.WriteAllLines(reportPath, reportLines);

        Console.WriteLine($"Summary report generated at: {reportPath}");
    }
}
