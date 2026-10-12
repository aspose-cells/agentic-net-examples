// Title: Generate a SmartArt shape count report for multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a list of .xlsx files with Aspose.Cells, iterates each worksheet, counts SmartArt shapes using ShapeType.SmartArt, and writes workbook name, worksheet name, and count to a new Excel report. | Modify the SmartArt report generator to add a summary row that shows the total number of SmartArt shapes per workbook. | Add functionality to export the SmartArt count report to a CSV file in addition to the Excel workbook, using Aspose.Cells or standard .NET I/O.
// Common Searches: how to count SmartArt objects in each worksheet with Aspose.Cells C# | Aspose.Cells enumerate shapes and get SmartArt count per sheet | create a summary Excel file that lists workbook and worksheet names with SmartArt counts | C# generate report of SmartArt shapes across multiple Excel workbooks using Aspose.Cells | detect SmartArt shapes in Excel files using Aspose.Cells ShapeType enumeration
// Tags: Aspose.Cells SmartArt shape counting | C# create SmartArt summary workbook | multiple workbook shape analysis Aspose.Cells | ShapeType.SmartArt detection C# | export SmartArt statistics CSV Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a predefined array of Excel file paths, opens each workbook with Aspose.Cells, iterates through every worksheet, counts shapes whose type is SmartArt (using the ShapeType enumeration), and records the workbook name, worksheet name, and SmartArt count. It then creates a new workbook named "SmartArt Report", writes headers and the collected data, auto‑fits columns, and saves the report as SmartArtReport.xlsx, handling missing files and runtime errors gracefully.
class SmartArtReportGenerator
{
    static void Main()
    {
        // List of workbook file paths to process
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            // Add more workbook paths as needed
        };

        // Prepare a list to hold report rows
        var reportRows = new List<(string WorkbookName, string WorksheetName, int SmartArtCount)>();

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Ensure the workbook file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook wb = new Workbook(filePath);

                // Get the workbook name (file name without path)
                string workbookName = Path.GetFileName(filePath);

                // Iterate through each worksheet
                foreach (Worksheet ws in wb.Worksheets)
                {
                    int smartArtCount = 0;

                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in ws.Shapes)
                    {
                        // Count shapes whose type string equals "SmartArt"
                        if (shape.Type.ToString() == "SmartArt")
                        {
                            smartArtCount++;
                        }
                    }

                    // Add the result to the report collection
                    reportRows.Add((workbookName, ws.Name, smartArtCount));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        try
        {
            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "SmartArt Report";

            // Write headers
            reportSheet.Cells[0, 0].PutValue("Workbook Name");
            reportSheet.Cells[0, 1].PutValue("Worksheet Name");
            reportSheet.Cells[0, 2].PutValue("SmartArt Count");

            // Write data rows
            int rowIndex = 1;
            foreach (var row in reportRows)
            {
                reportSheet.Cells[rowIndex, 0].PutValue(row.WorkbookName);
                reportSheet.Cells[rowIndex, 1].PutValue(row.WorksheetName);
                reportSheet.Cells[rowIndex, 2].PutValue(row.SmartArtCount);
                rowIndex++;
            }

            // Auto-fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            string reportPath = "SmartArtReport.xlsx";
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Report saved to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating report: {ex.Message}");
        }
    }
}
