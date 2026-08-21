// Title: C# – Merge Multiple Excel Workbooks into a Single PDF While Preserving Sheet Order with Aspose.Cells
// Description: Loads a collection of Excel files, uses the first workbook as the destination, calls Workbook.Combine to append the remaining worksheets in their original sequence, sets PdfSaveOptions.SheetSet to All, and saves the merged workbook as one PDF. Demonstrates Aspose.Cells for .NET conversion of all sheets to PDF with order retained.
// Keywords: Aspose.Cells | C# merge Excel workbooks | combine worksheets PDF | PdfSaveOptions SheetSet.All | export all sheets to PDF | preserve sheet order | Aspose.Cells Combine method | Excel to PDF conversion .NET
// Common Searches: Aspose.Cells combine workbooks C# | merge multiple Excel files into one PDF .NET | preserve worksheet order when exporting to PDF | export all Excel sheets to PDF using Aspose.Cells | C# code to create single PDF from several workbooks
// Developer Intent: Generate a single PDF that contains every worksheet from multiple Excel workbooks, keeping the original sheet sequence intact.
// Use Cases: Consolidate monthly financial workbooks into one PDF report with months appearing chronologically. | Combine departmental spreadsheets for an audit packet, ensuring each department's sheet order is maintained. | Create a printable PDF booklet from separate project Excel files, preserving the intended content flow across sheets.
// AI Prompts: Provide C# code using Aspose.Cells to merge a list of Excel workbooks and export all worksheets to one PDF while keeping the original sheet order. | Explain how PdfSaveOptions.SheetSet = SheetSet.All affects PDF output when converting a combined workbook with Aspose.Cells. | Show how to use Workbook.Combine in Aspose.Cells to append worksheets from multiple workbooks without losing their original sequence.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace MergeWorksheetsToPdf
{
    // Loads a collection of Excel files, uses the first workbook as the destination, calls Workbook.Combine to append the remaining worksheets in their original sequence, sets PdfSaveOptions.SheetSet to All, and saves the merged workbook as one PDF. Demonstrates Aspose.Cells for .NET conversion of all sheets to PDF with order retained.
    class Program
    {
        static void Main()
        {
            // Paths of the source Excel files (each may contain one or more worksheets)
            List<string> sourceFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Create the first workbook (will serve as the destination for merging)
            Workbook mergedWorkbook = new Workbook(sourceFiles[0]); // load first file

            // Combine the remaining workbooks preserving their original sheet order
            for (int i = 1; i < sourceFiles.Count; i++)
            {
                Workbook wb = new Workbook(sourceFiles[i]); // load next file
                mergedWorkbook.Combine(wb); // append its sheets after existing ones
            }

            // Prepare PDF save options to include all sheets in their original order
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                SheetSet = SheetSet.All // ensures every sheet is rendered in order
            };

            // Save the merged workbook as a single PDF file
            string outputPdf = "MergedOutput.pdf";
            mergedWorkbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"Merged PDF created successfully at: {outputPdf}");
        }
    }
}
