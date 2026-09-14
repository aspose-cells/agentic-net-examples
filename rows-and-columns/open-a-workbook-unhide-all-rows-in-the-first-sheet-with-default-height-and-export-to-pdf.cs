// Title: Unhide all rows in the first worksheet with default height and export the workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, uses Cells.UnhideRows to reveal every row in the first worksheet while keeping the default row height, and then saves the workbook as a PDF. | Show the exact sequence of Aspose.Cells API calls required to determine the total number of rows, unhide them from index 0, and generate a PDF output in a .NET application. | Demonstrate how to call Workbook.Save with SaveFormat.Pdf after changing row visibility in a worksheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# unhide all rows in first worksheet default height then convert to PDF | How to export an Excel file to PDF after making hidden rows visible with Aspose.Cells in .NET | C# code example for Cells.UnhideRows followed by Workbook.Save as PDF using Aspose.Cells
// Tags: cells.unhiderows default height aspose.cells c# | export worksheet to pdf aspose.cells c# | unhide all rows first worksheet aspose.cells | save workbook as pdf after row visibility change aspose.cells

using System;
using Aspose.Cells;

namespace AsposeCellsUnhideRowsAndExportPdf
{
    // // Loads an .xlsx workbook, unhides every row in the first sheet using the default height, and saves the result as a PDF file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Unhide all rows in the first sheet.
            // Use -1 for the height parameter to keep the default (auto‑fit) height.
            // The total number of rows is obtained from the worksheet's Rows collection.
            int totalRows = cells.Rows.Count;          // total rows in the sheet
            cells.UnhideRows(0, totalRows, -1);        // unhide from row 0 to the last row

            // Export the workbook to PDF.
            // The Save method automatically determines the format from the file extension.
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
