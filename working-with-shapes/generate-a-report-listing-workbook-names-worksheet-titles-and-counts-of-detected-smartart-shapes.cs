// Title: Generate a summary Excel file with workbook names, worksheet titles, and SmartArt shape counts using Aspose.Cells in C#
// AI Prompts: Write a C# console application that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet, counts shapes where IsSmartArt is true, and records the workbook name, worksheet title, and SmartArt count in a new Excel report. | Update the SmartArt counting utility to skip hidden worksheets, log any files that fail to load, and still produce the consolidated summary workbook. | Enhance the program to export the SmartArt summary as a CSV file and optionally add a column showing the total number of shapes per worksheet.
// Common Searches: how to count SmartArt objects in each worksheet with Aspose.Cells C# | C# Aspose.Cells generate report of SmartArt shapes across multiple Excel files | list worksheet names and SmartArt count using Aspose.Cells .NET | batch process Excel workbooks to summarize SmartArt shapes in C#
// Tags: Aspose.Cells SmartArt shape counting | C# Excel workbook summary generation | batch processing of .xlsx files with Aspose | detect SmartArt objects in worksheets .NET | export analysis results to CSV using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# console program that traverses a folder of .xlsx files, loads each workbook with Aspose.Cells, iterates through its worksheets, counts shapes flagged as SmartArt, and writes the workbook name, worksheet title, and SmartArt count into a new Excel summary workbook (or optionally a CSV).
class SmartArtReportGenerator
{
    static void Main()
    {
        try
        {
            // Folder containing the workbooks to analyze
            string sourceFolder = @"C:\Workbooks";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Output report workbook path
            string reportPath = @"C:\Reports\SmartArtReport.xlsx";

            // Ensure the report directory exists
            string? reportDir = Path.GetDirectoryName(reportPath);
            if (string.IsNullOrEmpty(reportDir))
            {
                Console.WriteLine("Invalid report path.");
                return;
            }
            if (!Directory.Exists(reportDir))
                Directory.CreateDirectory(reportDir);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "SmartArt Summary";

            // Write header row
            reportSheet.Cells[0, 0].PutValue("Workbook Name");
            reportSheet.Cells[0, 1].PutValue("Worksheet Title");
            reportSheet.Cells[0, 2].PutValue("SmartArt Count");

            int currentRow = 1; // Start after header

            // Get all Excel files in the source folder (including subfolders)
            string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.AllDirectories);

            foreach (string filePath in workbookFiles)
            {
                // Verify the file still exists before loading
                if (!File.Exists(filePath))
                    continue;

                Workbook wb;
                try
                {
                    wb = new Workbook(filePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook '{filePath}': {ex.Message}");
                    continue;
                }

                // Iterate through each worksheet
                foreach (Worksheet ws in wb.Worksheets)
                {
                    int smartArtCount = 0;

                    // Count SmartArt shapes in the worksheet
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Use IsSmartArt property for compatibility across versions
                        if (shape.IsSmartArt)
                            smartArtCount++;
                    }

                    // Write the data to the report sheet
                    reportSheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                    reportSheet.Cells[currentRow, 1].PutValue(ws.Name);
                    reportSheet.Cells[currentRow, 2].PutValue(smartArtCount);
                    currentRow++;
                }
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            try
            {
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
