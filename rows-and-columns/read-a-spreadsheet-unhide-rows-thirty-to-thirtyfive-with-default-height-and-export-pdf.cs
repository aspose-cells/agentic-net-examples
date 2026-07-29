// Title: C# – Unhide Rows 30‑35 (default height) and Convert Excel to PDF with Aspose.Cells
// Description: Load an Excel workbook, unhide rows 30‑35 using the UnhideRows method with a height of -1 (auto‑fit), and save the worksheet as a PDF file via Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# unhide rows | UnhideRows default height | Excel to PDF conversion .NET | Aspose.Cells SaveFormat.Pdf | auto‑fit row height Aspose
// Common Searches: Aspose.Cells unhide rows 30 to 35 C# | how to set default row height when unhiding Aspose.Cells | export Excel worksheet to PDF after unhiding rows | UnhideRows method height -1 example | C# convert hidden rows Excel to PDF
// Developer Intent: Unhide rows 30‑35 with their default height and generate a PDF from the updated Excel file.
// Use Cases: Prepare printable reports where previously hidden rows must appear at normal height before PDF creation. | Automate batch processing of workbooks that require specific rows to be visible and then converted to PDF for distribution. | Build a document‑generation service that ensures certain rows are auto‑fit and included in the final PDF output.
// AI Prompts: Generate C# code that uses Aspose.Cells to unhide rows 30‑35 with default height and save the workbook as a PDF. | Explain the parameters of Worksheet.Cells.UnhideRows, especially the use of -1 for automatic row height. | Show how to add error handling for loading an Excel file and exporting it to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Load an Excel workbook, unhide rows 30‑35 using the UnhideRows method with a height of -1 (auto‑fit), and save the worksheet as a PDF file via Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unhide rows 30 to 35 (zero‑based index 29, total 6 rows) with default (auto‑fit) height
            worksheet.Cells.UnhideRows(29, 6, -1);

            // Export the workbook to PDF format
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
