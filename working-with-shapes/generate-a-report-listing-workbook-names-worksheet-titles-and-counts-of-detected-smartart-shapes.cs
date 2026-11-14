// Title: Generate a worksheet‑wise SmartArt shape count report for multiple Excel workbooks with Aspose.Cells for .NET (C#)
// Common Searches: asp.net count smartart shapes in each worksheet using aspose.cells | c# create summary Excel file listing smartart count per sheet | batch process .xlsx files to report smartart objects with Aspose.Cells | how to enumerate smartart objects in worksheets via Aspose.Cells C# | generate smartart detection report for all workbooks in a folder | directory based Excel workbook processing for smartart counting using Aspose.Cells
// Tags: smartart shape enumeration Aspose.Cells C# | worksheet level smartart count batch processing | generate summary workbook Aspose.Cells | auto‑fit columns report Aspose.Cells | directory based Excel workbook analysis Aspose.Cells | smartart detection in Excel using Aspose.Cells API

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReportGenerator
{
    // Scans every .xlsx file in a specified folder, counts SmartArt shapes on each worksheet, and creates a new Excel report that lists the workbook name, worksheet title, and SmartArt count per sheet.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define the folder containing the workbooks to analyze.
                // Adjust this path as needed.
                string sourceFolder = @"C:\Workbooks";

                // Verify the source folder exists.
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                    return;
                }

                // Get all Excel files in the folder (top level only).
                string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

                // Prepare a list to hold report rows.
                var reportRows = new List<(string WorkbookName, string WorksheetName, int SmartArtCount)>();

                // Iterate through each workbook file.
                foreach (string filePath in workbookFiles)
                {
                    // Ensure the file exists before loading.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    // Load the workbook.
                    Workbook wb = new Workbook(filePath);

                    // Iterate through each worksheet in the workbook.
                    foreach (Worksheet ws in wb.Worksheets)
                    {
                        // Count SmartArt shapes on the current worksheet.
                        int smartArtCount = 0;
                        foreach (Shape shape in ws.Shapes)
                        {
                            // Use the IsSmartArt property to identify SmartArt shapes.
                            if (shape.IsSmartArt)
                            {
                                smartArtCount++;
                            }
                        }

                        // Add the information to the report list.
                        reportRows.Add((Path.GetFileName(filePath), ws.Name, smartArtCount));
                    }
                }

                // Create a new workbook for the report.
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "SmartArt Report";

                // Write header row.
                reportSheet.Cells[0, 0].PutValue("Workbook Name");
                reportSheet.Cells[0, 1].PutValue("Worksheet Title");
                reportSheet.Cells[0, 2].PutValue("SmartArt Count");

                // Populate the report with collected data.
                int rowIndex = 1;
                foreach (var row in reportRows)
                {
                    reportSheet.Cells[rowIndex, 0].PutValue(row.WorkbookName);
                    reportSheet.Cells[rowIndex, 1].PutValue(row.WorksheetName);
                    reportSheet.Cells[rowIndex, 2].PutValue(row.SmartArtCount);
                    rowIndex++;
                }

                // Auto-fit columns for better readability.
                reportSheet.AutoFitColumns();

                // Save the report workbook.
                string reportPath = Path.Combine(sourceFolder, "SmartArtReport.xlsx");
                reportWorkbook.Save(reportPath, SaveFormat.Xlsx);

                Console.WriteLine($"Report generated successfully at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
