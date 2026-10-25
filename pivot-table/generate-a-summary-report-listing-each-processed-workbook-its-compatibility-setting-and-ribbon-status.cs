// Title: Generate a consolidated Excel summary of workbook Compatibility setting and ribbon enablement using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read workbook Compatibility property from multiple Excel files | How to list IsRibbonsEnabled status for a set of workbooks with Aspose.Cells | Create summary worksheet with workbook settings using Aspose.Cells in .NET | Aspose.Cells reflection to get Compatibility and ribbon flag from Excel workbooks | Save aggregated workbook settings report as XLSX with Aspose.Cells
// Tags: Aspose.Cells read workbook compatibility | Aspose.Cells extract ribbon enabled flag | Aspose.Cells generate summary worksheet | Aspose.Cells save workbook as xlsx | C# reflection Aspose.Cells Settings properties

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSummaryReport
{
    // The C# program creates a new workbook named "Summary", adds a header row, and iterates over a list of Excel file paths. For each existing file it loads the workbook with Aspose.Cells, uses reflection to read the Settings.Compatibility and Settings.IsRibbonsEnabled properties (when available), records the file name, compatibility value, and ribbon status into the summary sheet, and finally saves the aggregated report as "SummaryReport.xlsx".
    class Program
    {
        static void Main(string[] args)
        {
            // Define the list of workbook file paths to process.
            // Adjust the paths as needed for your environment.
            string[] workbookFiles = new string[]
            {
                @"C:\Workbooks\Book1.xlsx",
                @"C:\Workbooks\Book2.xlsx",
                @"C:\Workbooks\Book3.xlsx"
            };

            // Create a new workbook that will hold the summary report.
            Workbook summaryWorkbook = new Workbook(); // create-workbook rule
            Worksheet summarySheet = summaryWorkbook.Worksheets[0];
            summarySheet.Name = "Summary";

            // Write header row.
            summarySheet.Cells[0, 0].PutValue("Workbook");
            summarySheet.Cells[0, 1].PutValue("Compatibility");
            summarySheet.Cells[0, 2].PutValue("Ribbon Enabled");

            // Process each workbook and collect required information.
            for (int i = 0; i < workbookFiles.Length; i++)
            {
                string filePath = workbookFiles[i];

                // Verify that the file exists before attempting to load it.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook.
                    Workbook wb = new Workbook(filePath); // load-workbook rule

                    // Retrieve compatibility setting via reflection (covers different API versions).
                    string compatibility = "N/A";
                    var compProp = wb.Settings.GetType().GetProperty("Compatibility");
                    if (compProp != null)
                    {
                        object compValue = compProp.GetValue(wb.Settings);
                        compatibility = compValue?.ToString() ?? "N/A";
                    }

                    // Retrieve ribbon status via reflection (covers different API versions).
                    string ribbonStatus = "N/A";
                    var ribbonProp = wb.Settings.GetType().GetProperty("IsRibbonsEnabled");
                    if (ribbonProp != null)
                    {
                        object ribbonValue = ribbonProp.GetValue(wb.Settings);
                        if (ribbonValue is bool isEnabled)
                            ribbonStatus = isEnabled ? "Enabled" : "Disabled";
                    }

                    // Write the information into the summary sheet.
                    int row = i + 1; // data starts after header
                    summarySheet.Cells[row, 0].PutValue(Path.GetFileName(filePath));
                    summarySheet.Cells[row, 1].PutValue(compatibility);
                    summarySheet.Cells[row, 2].PutValue(ribbonStatus);
                }
                catch (Exception ex)
                {
                    // Log the error and continue processing remaining workbooks.
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            // Save the summary report workbook.
            string outputPath = @"C:\Workbooks\SummaryReport.xlsx";
            try
            {
                summaryWorkbook.Save(outputPath, SaveFormat.Xlsx); // save-workbook rule
                Console.WriteLine("Summary report generated at: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save summary report: {ex.Message}");
            }
        }
    }
}
