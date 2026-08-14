// Title: Batch SmartArt Shape Count Report for Excel Workbooks using Aspose.Cells for .NET (C#)
// Description: A C# utility that scans every .xlsx file in a given folder, loads each workbook with Aspose.Cells, iterates through all worksheets, counts shapes flagged as SmartArt, and writes a summary workbook (SmartArtReport.xlsx) containing the source workbook name, worksheet title, and SmartArt count per sheet.
// Keywords: Aspose.Cells SmartArt count | C# batch Excel shape analysis | count SmartArt per worksheet | generate Excel report with Aspose.Cells | automate SmartArt inventory .NET | Excel shape detection C# | bulk workbook processing Aspose
// Common Searches: how to count SmartArt shapes in Excel using Aspose.Cells | C# program to list SmartArt objects across multiple workbooks | create summary sheet of SmartArt counts with Aspose.Cells | batch process Excel files for SmartArt statistics | Aspose.Cells shape enumeration example
// Developer Intent: Produce a consolidated Excel file that lists each source workbook, its worksheets, and the number of SmartArt shapes found on each worksheet.
// Use Cases: Audit a corporate template library to ensure design consistency of SmartArt usage. | Compile an inventory of SmartArt elements before migrating Excel assets to a new platform. | Monitor SmartArt density in generated reports to maintain performance and file size limits.
// AI Prompts: Write C# code with Aspose.Cells that counts SmartArt shapes per worksheet and outputs the data to a new Excel summary file. | Extend the program to add a column showing the total SmartArt count for each workbook. | Suggest robust error‑handling patterns for missing files, corrupted workbooks, and inaccessible folders when counting SmartArt shapes.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtReportGenerator
{
    // A C# utility that scans every .xlsx file in a given folder, loads each workbook with Aspose.Cells, iterates through all worksheets, counts shapes flagged as SmartArt, and writes a summary workbook (SmartArtReport.xlsx) containing the source workbook name, worksheet title, and SmartArt count per sheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Folder containing the workbooks to be analyzed
                string inputFolder = @"InputWorkbooks";

                // Verify the input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder not found: {inputFolder}");
                    return;
                }

                // Get all Excel files in the folder
                string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

                // Create a new workbook that will hold the report
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];

                // Write header row
                reportSheet.Cells[0, 0].PutValue("Workbook");
                reportSheet.Cells[0, 1].PutValue("Worksheet");
                reportSheet.Cells[0, 2].PutValue("SmartArt Count");

                int reportRow = 1; // Start writing data from the second row

                // Process each workbook
                foreach (string filePath in workbookFiles)
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found, skipping: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook wb = new Workbook(filePath);
                        string workbookName = Path.GetFileName(filePath);

                        // Iterate through all worksheets
                        foreach (Worksheet ws in wb.Worksheets)
                        {
                            int smartArtCount = 0;

                            // Count shapes that are SmartArt
                            foreach (Shape shape in ws.Shapes)
                            {
                                if (shape.IsSmartArt)
                                {
                                    smartArtCount++;
                                }
                            }

                            // Write the information to the report sheet
                            reportSheet.Cells[reportRow, 0].PutValue(workbookName);
                            reportSheet.Cells[reportRow, 1].PutValue(ws.Name);
                            reportSheet.Cells[reportRow, 2].PutValue(smartArtCount);
                            reportRow++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                // Save the report workbook
                string reportPath = @"SmartArtReport.xlsx";
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report generated: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
