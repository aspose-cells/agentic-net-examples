// Title: Hide rows 10‑20, unhide rows 15‑18 with auto‑fit, and export the worksheet to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Hide rows 10 through 20 in a worksheet, then unhide rows 15‑18 with auto‑fit height, and save the workbook as a PDF using Aspose.Cells in C#. | Load an Excel file, apply HideRows and UnhideRows methods to adjust row visibility, and generate a PDF output with Aspose.Cells for .NET. | Programmatically change row visibility ranges and export the result to PDF using Aspose.Cells' HideRows, UnhideRows, and SaveFormat.Pdf in C#.
// Common Searches: Aspose.Cells C# hide rows 10-20 then unhide rows 15-18 and export to PDF | how to auto‑fit row height after unhiding rows using Aspose.Cells .NET | C# code to hide a range of rows and save the workbook as PDF with Aspose.Cells
// Tags: HideRows method Aspose.Cells C# | UnhideRows with auto‑fit height Aspose.Cells | Export worksheet to PDF Aspose.Cells | Row visibility manipulation Excel Aspose.Cells | PDF conversion after row visibility changes Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRowVisibilityDemo
{
    // Loads input.xlsx, hides rows 10‑20, unhides rows 15‑18 with auto‑fit height, and saves the result as output.pdf using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Hide rows 10 to 20 (zero‑based index: start at 9, total 11 rows)
            cells.HideRows(9, 11);

            // Unhide rows 15 to 18 (zero‑based index: start at 14, total 4 rows)
            // Height = -1 means auto‑fit the row height after unhiding
            cells.UnhideRows(14, 4, -1);

            // Save the modified workbook as PDF
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
