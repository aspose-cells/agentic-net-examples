// Title: C# Aspose.Cells – Generate a SmartArt Count Report for Multiple Excel Workbooks
// Description: A console application that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet and its Shapes collection, counts shapes where Shape.IsSmartArt is true, and writes the workbook name, worksheet title, and SmartArt total to a new report workbook (SmartArtReport.xlsx). Ideal for batch analysis and documentation of SmartArt usage in Excel files on Windows/.NET platforms.
// Keywords: Aspose.Cells | C# SmartArt count | Excel shape enumeration | batch workbook analysis | generate SmartArt report | Shape.IsSmartArt | .NET Excel automation | folder scan Excel files | report workbook creation
// Common Searches: count SmartArt shapes in each worksheet using Aspose.Cells | C# generate SmartArt summary across multiple Excel files | Aspose.Cells iterate shapes and detect SmartArt | export SmartArt totals to a new Excel workbook | batch process Excel files for SmartArt statistics
// Developer Intent: Create an automated Excel report that lists every workbook and worksheet together with the number of SmartArt objects it contains, using Aspose.Cells for .NET.
// Use Cases: Audit a directory of Excel workbooks to identify worksheets that include SmartArt for compliance or documentation. | Provide stakeholders with a concise overview of SmartArt usage across project reports. | Integrate into a CI pipeline to flag worksheets that exceed expected SmartArt counts.
// AI Prompts: Write C# code with Aspose.Cells that counts SmartArt shapes in all worksheets of multiple workbooks and saves the results to a new Excel file. | Explain the behavior of Shape.IsSmartArt and how to handle scenarios where the property is unavailable in older Aspose.Cells versions. | Suggest enhancements for the SmartArt report, such as adding hyperlinks to source worksheets, summarizing totals per workbook, or exporting to CSV.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReportGenerator
{
    // A console application that scans a folder for *.xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet and its Shapes collection, counts shapes where Shape.IsSmartArt is true, and writes the workbook name, worksheet title, and SmartArt total to a new report workbook (SmartArtReport.xlsx). Ideal for batch analysis and documentation of SmartArt usage in Excel files on Windows/.NET platforms.
    class Program
    {
        // Simple DTO to hold report rows
        class ReportRow
        {
            public string WorkbookName { get; set; }
            public string WorksheetName { get; set; }
            public int SmartArtCount { get; set; }
        }

        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Define the Excel files to be processed.
            //    Adjust the folder path and file filter as needed.
            // -----------------------------------------------------------------
            string folderPath = @"C:\ExcelFiles";
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

            // -----------------------------------------------------------------
            // 2. Collect report data.
            // -----------------------------------------------------------------
            List<ReportRow> reportData = new List<ReportRow>();

            foreach (string filePath in excelFiles)
            {
                // Load workbook (uses the provided load rule)
                Workbook workbook = new Workbook(filePath);

                // Iterate through worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int smartArtCount = 0;

                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Use the Shape.IsSmartArt property (rule exists)
                        if (shape.IsSmartArt)
                        {
                            smartArtCount++;
                        }
                    }

                    // Store the result for this worksheet
                    reportData.Add(new ReportRow
                    {
                        WorkbookName = Path.GetFileName(filePath),
                        WorksheetName = sheet.Name,
                        SmartArtCount = smartArtCount
                    });
                }
            }

            // -----------------------------------------------------------------
            // 3. Create a new workbook to hold the report (uses the create rule)
            // -----------------------------------------------------------------
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "SmartArt Report";

            // Write header
            reportSheet.Cells["A1"].PutValue("Workbook Name");
            reportSheet.Cells["B1"].PutValue("Worksheet Title");
            reportSheet.Cells["C1"].PutValue("SmartArt Count");

            // Write data rows
            int currentRow = 1; // zero‑based index; row 1 is the second row (after header)
            foreach (ReportRow row in reportData)
            {
                reportSheet.Cells[currentRow, 0].PutValue(row.WorkbookName);
                reportSheet.Cells[currentRow, 1].PutValue(row.WorksheetName);
                reportSheet.Cells[currentRow, 2].PutValue(row.SmartArtCount);
                currentRow++;
            }

            // -----------------------------------------------------------------
            // 4. Save the report workbook (uses the save rule)
            // -----------------------------------------------------------------
            string reportPath = Path.Combine(folderPath, "SmartArtReport.xlsx");
            reportWorkbook.Save(reportPath);

            Console.WriteLine($"Report generated successfully at: {reportPath}");
        }
    }
}
