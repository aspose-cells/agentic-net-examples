// Title: C# – Export Worksheets 1 and 3 to PDF with Aspose.Cells LightCells API
// Description: Demonstrates how to load an Excel workbook, select the first and third worksheets (using zero‑based indexes and a fallback for missing sheets), configure PdfSaveOptions with a SheetSet, and save only those sheets as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells LightCells API | C# export selected worksheets to PDF | PdfSaveOptions SheetSet example | load workbook specific sheets PDF | Aspose.Cells sheet selection | Excel to PDF selective conversion
// Common Searches: Aspose.Cells export only sheet 1 and 3 to PDF C# | PdfSaveOptions SheetSet usage example | How to save selected worksheets as PDF with Aspose.Cells | C# LightCells API save specific sheets to PDF | Select worksheets for PDF export Aspose.Cells
// Developer Intent: Generate a PDF that contains only the first and third worksheets of an Excel file.
// Use Cases: Create a concise PDF report that includes only summary and data sheets from a multi‑sheet workbook. | Reduce file size for client‑facing PDFs by exporting only relevant worksheets. | Automate batch conversions where compliance documents require specific sheets from each workbook.
// AI Prompts: Write C# code using Aspose.Cells to export worksheets 2 and 4 to a PDF with custom page orientation. | Show how to build a dynamic SheetSet list for PdfSaveOptions when the number of target worksheets varies. | Explain strategies for handling missing worksheet indexes while configuring SheetSet for PDF export.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLightCellsExample
{
    // Demonstrates how to load an Excel workbook, select the first and third worksheets (using zero‑based indexes and a fallback for missing sheets), configure PdfSaveOptions with a SheetSet, and save only those sheets as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourceFile = "input.xlsx";

                // Ensure the source file exists; if not, create a simple workbook with three sheets
                if (!File.Exists(sourceFile))
                {
                    var tempWb = new Workbook();
                    // Create three worksheets with sample data
                    for (int i = 0; i < 3; i++)
                    {
                        Worksheet ws = tempWb.Worksheets[i];
                        ws.Name = $"Sheet{i + 1}";
                        ws.Cells["A1"].PutValue($"Data in {ws.Name}");
                        if (i > 0)
                            tempWb.Worksheets.Add();
                    }
                    tempWb.Save(sourceFile);
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Determine which sheet indexes are valid (zero‑based)
                List<int> validIndexes = new List<int>();
                if (workbook.Worksheets.Count > 0) validIndexes.Add(0);          // first sheet
                if (workbook.Worksheets.Count > 2) validIndexes.Add(2);          // third sheet
                else if (workbook.Worksheets.Count > 1) validIndexes.Add(1);    // fallback to second sheet if third absent

                // Create PDF save options and specify the selected sheets
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = new SheetSet(validIndexes.ToArray())
                };

                // Save the selected sheets to a PDF file
                workbook.Save("output.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
