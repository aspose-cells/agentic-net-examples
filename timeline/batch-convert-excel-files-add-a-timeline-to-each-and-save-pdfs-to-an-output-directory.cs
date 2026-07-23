// Title: Batch Convert Excel to PDF with PivotTable Timeline in C# using Aspose.Cells
// Description: A C# console app that scans an input folder for .xlsx files, adds a pivot table and a linked timeline to each workbook, saves a temporary copy, exports the workbook to PDF in an output folder, and cleans up temporary files while handling per‑file errors.
// Keywords: Aspose.Cells | C# | .NET | batch convert Excel to PDF | pivot table | timeline | folder processing | automated PDF export | workbook automation | temporary file cleanup
// Common Searches: Aspose.Cells add timeline to pivot table C# | batch convert multiple Excel files to PDF using Aspose.Cells | C# code to export Excel workbooks with pivot tables to PDF | automate Excel to PDF conversion with timeline feature | how to process all .xlsx files in a directory with Aspose.Cells
// Developer Intent: Write a C# program that iterates through Excel files, inserts a pivot table with a date timeline, and saves each workbook as a PDF.
// Use Cases: Generate PDF sales dashboards by adding a date timeline to pivot tables for each monthly Excel report. | Automate archival of incoming financial worksheets, applying a timeline filter before converting them to searchable PDFs. | Prepare compliance packages by batch‑processing Excel logs, inserting pivot tables with timelines, and exporting PDFs for audit review.
// AI Prompts: Create C# code that loads every .xlsx in a folder, adds a pivot table and a linked timeline, then saves the workbook as PDF using Aspose.Cells. | Suggest robust error handling and temporary file management for batch Excel‑to‑PDF conversion with Aspose.Cells. | Explain how to customize the position, size, and style of a timeline added to a worksheet via Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

// A C# console app that scans an input folder for .xlsx files, adds a pivot table and a linked timeline to each workbook, saves a temporary copy, exports the workbook to PDF in an output folder, and cleans up temporary files while handling per‑file errors.
class BatchConvertWithTimeline
{
    static void Main()
    {
        // Input and output directories
        string inputDir = @"C:\InputExcels";
        string outputDir = @"C:\OutputPdfs";

        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory does not exist: {inputDir}");
            return;
        }

        Directory.CreateDirectory(outputDir);

        // Process each .xlsx file
        foreach (string excelFile in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            try
            {
                if (!File.Exists(excelFile))
                {
                    Console.WriteLine($"File not found: {excelFile}");
                    continue;
                }

                // Load workbook
                Workbook wb = new Workbook(excelFile);
                Worksheet sheet = wb.Worksheets[0];

                // Add sample data if sheet is empty
                if (sheet.Cells.MaxDataRow == 0 && sheet.Cells.MaxDataColumn == 0)
                {
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Date");
                    sheet.Cells["C1"].PutValue("Amount");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["B2"].PutValue(DateTime.Now);
                    sheet.Cells["C2"].PutValue(100);
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["B3"].PutValue(DateTime.Now.AddDays(1));
                    sheet.Cells["C3"].PutValue(200);
                }

                // Determine used range for pivot table
                int firstRow = sheet.Cells.MinRow;
                int firstCol = sheet.Cells.MinColumn;
                int lastRow = sheet.Cells.MaxDataRow;
                int lastCol = sheet.Cells.MaxDataColumn;
                string dataRange = CellsHelper.CellIndexToName(firstRow, firstCol) + ":" +
                                   CellsHelper.CellIndexToName(lastRow, lastCol);

                // Add pivot table
                int pivotIndex = sheet.PivotTables.Add(dataRange, "E3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                pivot.AddFieldToArea(PivotFieldType.Row, 1);   // Date
                pivot.AddFieldToArea(PivotFieldType.Data, 2);  // Amount

                // Add timeline linked to Date field
                sheet.Timelines.Add(pivot, "G1", "Date");

                // Save temporary workbook
                string tempPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(excelFile) + "_temp.xlsx");
                wb.Save(tempPath, SaveFormat.Xlsx);

                // Convert to PDF
                string pdfPath = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(excelFile) + ".pdf");
                wb.Save(pdfPath, SaveFormat.Pdf);

                // Clean up temporary file
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }

                Console.WriteLine($"Converted '{excelFile}' to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{excelFile}': {ex.Message}");
            }
        }
    }
}
