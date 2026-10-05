// Title: Auto-fit rows that contain merged cells before exporting an Excel worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Create C# code that enumerates all merged cell areas in a worksheet, calls worksheet.AutoFitRow for every row in those areas, and then saves the workbook as a PDF. | Write a reusable C# method with Aspose.Cells that adjusts the height of rows involved in merged ranges and returns a workbook ready for PDF conversion.
// Common Searches: c# asp.net auto fit rows with merged cells before pdf export using Aspose.Cells | how to adjust row height for merged cells when converting Excel to PDF with Aspose.Cells | Aspose.Cells merged cell row height issue in PDF output | auto fit rows for merged ranges Aspose.Cells .NET example | worksheet.AutoFitRow merged cells PDF conversion problem solution
// Tags: auto-fit rows merged cells Aspose.Cells | merged cell row height PDF export .NET | worksheet.AutoFitRow merged ranges | Aspose.Cells adjust row height for merged areas | Excel to PDF merged cells handling

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, detects every merged cell range, auto-fits each row that participates in those ranges to ensure proper display, and then saves the result as a PDF, including basic error handling for missing files.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve merged cell areas
                CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

                // Auto‑fit rows that participate in merged ranges
                foreach (CellArea area in mergedAreas)
                {
                    int firstRow = area.StartRow;
                    int lastRow = area.EndRow;

                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        worksheet.AutoFitRow(row);
                    }
                }

                // Save the workbook as PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
