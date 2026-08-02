// Title: Set A4 Paper Size for Every Worksheet in an Aspose.Cells Workbook (C# foreach Loop)
// Description: Creates a Workbook, adds extra sheets, then uses a foreach loop to set each worksheet's PageSetup.PaperSize to PaperSizeType.PaperA4 before saving the file as AllSheets_A4PaperSize.xlsx.
// Keywords: Aspose.Cells | C# | set paper size | A4 | foreach loop | PageSetup | worksheet print settings | Workbook API | Aspose.Cells tutorial
// Common Searches: Aspose.Cells set A4 paper size for all worksheets | C# foreach set worksheet page size Aspose.Cells | How to change print paper size in an Aspose.Cells workbook | Apply same page setup to multiple sheets Aspose.Cells | Set default paper size for new workbook using Aspose.Cells
// Developer Intent: Apply A4 paper size to every worksheet in a workbook using Aspose.Cells and a foreach loop.
// Use Cases: Standardize the print layout of multi‑sheet reports before exporting to PDF. | Ensure consistent A4 dimensions when batch‑printing a generated workbook. | Prepare a workbook for distribution where each sheet must conform to A4 page size.
// AI Prompts: Generate C# code that sets the paper size to Letter for all worksheets in an Aspose.Cells workbook using a foreach loop. | Show how to change the page orientation to landscape for each worksheet in an Aspose.Cells workbook. | Provide an example that applies A4 paper size and 1‑inch margins to every sheet in a workbook with Aspose.Cells. | Create a script that sets both A4 paper size and a custom header/footer for all worksheets in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    // Creates a Workbook, adds extra sheets, then uses a foreach loop to set each worksheet's PageSetup.PaperSize to PaperSizeType.PaperA4 before saving the file as AllSheets_A4PaperSize.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets for demonstration (optional)
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Iterate through all worksheets and set the paper size to A4
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // PageSetup.PaperSize controls the print paper size for the sheet
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }

            // Save the workbook to an XLSX file
            workbook.Save("AllSheets_A4PaperSize.xlsx", SaveFormat.Xlsx);
        }
    }
}
