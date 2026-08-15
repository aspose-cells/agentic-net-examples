// Title: Aspose.Cells .NET – Console report of workbook CheckCompatibility flag and RibbonXml presence
// Description: A C# console utility that iterates over a collection of Excel files, loads each workbook with Aspose.Cells, reads the Workbook.Settings.CheckCompatibility property, checks whether the RibbonXml string is defined, and prints a formatted table showing the file path, compatibility mode, and ribbon status. The program gracefully handles missing files and load exceptions, making it ideal for batch audits or automated reporting.
// Keywords: Aspose.Cells | .NET | C# | CheckCompatibility | RibbonXml | Excel workbook audit | batch workbook processing | console report | Excel compatibility mode | custom ribbon detection | automation script
// Common Searches: Aspose.Cells get CheckCompatibility value | how to detect RibbonXml in Excel file using Aspose.Cells | C# console program list workbook settings | batch check Excel compatibility mode Aspose.Cells | report ribbon XML presence in workbooks
// Developer Intent: Generate a concise summary that lists each processed workbook, its compatibility setting, and whether RibbonXml is defined, using Aspose.Cells in a .NET console application.
// Use Cases: Run a nightly compliance audit that records the compatibility mode of all corporate workbooks. | Verify that custom ribbon XML has been applied to macro‑enabled files before distribution. | Export workbook paths, compatibility flags, and ribbon status to CSV for downstream analytics.
// AI Prompts: Create C# code that reads CheckCompatibility and RibbonXml for each workbook and writes the results to a CSV file with Aspose.Cells. | Enhance the console program's error handling to log full exception details to a log file. | Show how to filter the workbook list to only those with RibbonXml set and output their file names.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSummaryReport
{
    // A C# console utility that iterates over a collection of Excel files, loads each workbook with Aspose.Cells, reads the Workbook.Settings.CheckCompatibility property, checks whether the RibbonXml string is defined, and prints a formatted table showing the file path, compatibility mode, and ribbon status. The program gracefully handles missing files and load exceptions, making it ideal for batch audits or automated reporting.
    class Program
    {
        static void Main()
        {
            // List of workbook file paths to process
            var workbookPaths = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsm",
                "Workbook3.xls"
                // Add more paths as needed
            };

            // Header for the summary report
            Console.WriteLine("Processed Workbook Summary");
            Console.WriteLine("---------------------------");
            Console.WriteLine("{0,-30} {1,-15} {2}", "File Path", "CheckCompatibility", "RibbonXml Set");

            // Process each workbook
            foreach (var path in workbookPaths)
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(path))
                {
                    Console.WriteLine("{0,-30} {1,-15} {2}", path, "N/A", "File not found");
                    continue;
                }

                try
                {
                    // Load the workbook from the file path
                    Workbook workbook = new Workbook(path);

                    // Retrieve the compatibility setting from WorkbookSettings
                    bool checkCompatibility = workbook.Settings.CheckCompatibility;

                    // Determine whether RibbonXml is set (non‑null and non‑empty)
                    bool ribbonSet = !string.IsNullOrEmpty(workbook.RibbonXml);

                    // Output the information for this workbook
                    Console.WriteLine("{0,-30} {1,-15} {2}",
                        path,
                        checkCompatibility,
                        ribbonSet ? "Yes" : "No");
                }
                catch (Exception ex)
                {
                    // Handle any unexpected errors during loading/processing
                    Console.WriteLine("{0,-30} {1,-15} {2}", path, "Error", ex.Message);
                }
            }

            // Keep console window open if needed
            Console.WriteLine("\nSummary report completed.");
        }
    }
}
