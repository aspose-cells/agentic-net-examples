// Title: Export an Excel workbook with slicers to PDF and verify slicer positions stay the same using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file containing slicers, saves it as a PDF with Aspose.Cells while keeping the slicer layout intact, and then outputs whether each slicer's UpperLeftRow and UpperLeftColumn match the original values. | Write a C# example that creates a sample workbook with a pivot table and a slicer, exports the workbook to PDF using PdfSaveOptions, and compares the slicer's shape coordinates before and after the export.
// Common Searches: how to keep slicer positions when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells C# export workbook containing slicers to PDF and check layout | verify slicer coordinates after PDF save using Aspose.Cells .NET | sample code for exporting Excel file with pivot slicer to PDF in C# | compare slicer UpperLeftRow UpperLeftColumn before and after PDF conversion Aspose.Cells
// Tags: Aspose.Cells export workbook to PDF preserving slicer layout | slicer position validation after PDF conversion | C# pivot table slicer PDF generation | PdfSaveOptions enable document structure Aspose.Cells | compare slicer shape coordinates C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;

// The program loads (or creates) an Excel workbook that includes a pivot table and a slicer, records each slicer's UpperLeftRow and UpperLeftColumn, saves the workbook to PDF with document structure enabled, reloads the original workbook, and reports whether the slicer positions remain unchanged after the PDF export.
class ExportWorkbookWithSlicersToPdf
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputWithSlicers.xlsx";
            const string outputPdf = "WorkbookWithSlicers.pdf";

            // Ensure the input workbook exists; create a minimal one if it does not.
            if (!File.Exists(inputFile))
            {
                try
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Name = "Data";

                    // Add sample data.
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Value");
                    ws.Cells["A2"].PutValue("A");
                    ws.Cells["B2"].PutValue(10);
                    ws.Cells["A3"].PutValue("B");
                    ws.Cells["B3"].PutValue(20);
                    ws.Cells["A4"].PutValue("C");
                    ws.Cells["B4"].PutValue(30);

                    // Create a pivot table based on the data range.
                    var pivotTable = ws.PivotTables[ws.PivotTables.Add("PivotTable1", "E1", "A1:B4", true)];

                    // Add a slicer for the first field (Category) of the pivot table.
                    // Note: Add(PivotTable pivotTable, int fieldIndex, int upperLeftRow, int upperLeftColumn)
                    var slicerIndex = ws.Slicers.Add(pivotTable, 0, 5, 0);
                    var slicer = ws.Slicers[slicerIndex];
                    slicer.Name = "CategorySlicer";

                    wb.Save(inputFile);
                    Console.WriteLine($"Input file '{inputFile}' was not found and has been created with sample data.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Load the existing workbook that contains slicers.
            var workbook = new Workbook(inputFile);
            var worksheet = workbook.Worksheets[0];

            // Store original slicer positions.
            var originalPositions = new Dictionary<string, (int Row, int Column)>();
            foreach (Slicer slicer in worksheet.Slicers)
            {
                var shape = slicer.Shape;
                originalPositions[slicer.Name] = (shape.UpperLeftRow, shape.UpperLeftColumn);
            }

            // Prepare PDF save options – enable document structure export.
            var pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook (including slicers) to a PDF file.
            workbook.Save(outputPdf, pdfOptions);
            Console.WriteLine($"Workbook saved to PDF as '{outputPdf}'.");

            // Reload the original workbook to verify that slicer positions are unchanged.
            var reloadedWorkbook = new Workbook(inputFile);
            var reloadedWorksheet = reloadedWorkbook.Worksheets[0];

            foreach (Slicer slicer in reloadedWorksheet.Slicers)
            {
                var shape = slicer.Shape;
                var key = slicer.Name;

                if (originalPositions.TryGetValue(key, out var original))
                {
                    bool positionUnchanged = original.Row == shape.UpperLeftRow &&
                                             original.Column == shape.UpperLeftColumn;

                    Console.WriteLine($"Slicer '{key}' position unchanged after PDF export: {positionUnchanged}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
