// Title: C# – Apply A4 Paper Size to Every Worksheet in an Aspose.Cells Workbook Using a foreach Loop
// Description: Creates or loads a workbook, adds sample sheets, iterates through all worksheets with a foreach loop, sets each sheet's PageSetup.PaperSize to PaperA4, and saves the file.
// Keywords: Aspose.Cells | C# | set paper size | A4 | foreach loop | worksheet page setup | Workbook PageSetup | multiple sheets | print layout | PDF conversion
// Common Searches: Aspose.Cells set A4 paper size for all worksheets | C# foreach loop change page setup in Aspose.Cells | apply same paper size to every sheet in a workbook | Aspose.Cells C# set page size for multiple worksheets | batch update worksheet page settings Aspose.Cells
// Developer Intent: Programmatically set the paper size of every worksheet in a workbook to A4.
// Use Cases: Standardize print layout for multi‑sheet reports before exporting to PDF or XPS. | Ensure consistent page dimensions when generating batch‑printed workbooks. | Prepare a workbook with uniform A4 pages for automated document workflows.
// AI Prompts: Generate C# code that changes the orientation to landscape for all worksheets in an Aspose.Cells workbook using a foreach loop. | Show how to set custom margins, header, and footer on every sheet of an Aspose.Cells workbook with C#. | Provide an example that applies the same page scaling factor to each worksheet in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates or loads a workbook, adds sample sheets, iterates through all worksheets with a foreach loop, sets each sheet's PageSetup.PaperSize to PaperA4, and saves the file.
class SetPaperSizeForAllSheets
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add additional worksheets for demonstration purposes
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Apply A4 paper size to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
        }

        // Save the workbook with the updated settings
        workbook.Save("AllSheetsA4.xlsx", SaveFormat.Xlsx);
    }
}
