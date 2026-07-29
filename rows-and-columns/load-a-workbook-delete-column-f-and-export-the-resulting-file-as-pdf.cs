// Title: Delete Column F in an Excel Workbook and Export to PDF using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file with Aspose.Cells, removes the sixth column (column F) from the first worksheet, and saves the modified workbook directly as a PDF document.
// Keywords: Aspose.Cells | C# delete column | remove Excel column | export to PDF | SaveFormat.Pdf | column F | worksheet column deletion | Excel to PDF conversion .NET | Aspose.Cells API | DeleteColumn method
// Common Searches: Aspose.Cells delete column C# | How to remove a column from Excel and save as PDF in .NET | C# convert Excel to PDF after column removal | Delete specific column using Aspose.Cells | Save modified workbook as PDF Aspose.Cells
// Developer Intent: Remove a specific column from an Excel sheet and generate a PDF version of the updated workbook.
// Use Cases: Prepare client‑ready PDF reports by omitting sensitive or irrelevant columns. | Create streamlined PDFs for dashboards where certain data fields are unnecessary. | Automate bulk processing to strip unwanted columns from many workbooks before archiving them as PDFs.
// AI Prompts: Generate C# code with Aspose.Cells to delete column H and export the worksheet to a landscape‑oriented PDF. | Explain how to delete multiple adjacent columns while preserving cell styles, then save the workbook as a PDF using Aspose.Cells. | Show how to locate a column by its header text, delete it, and convert the sheet to PDF in a .NET application.

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, removes the sixth column (column F) from the first worksheet, and saves the modified workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Delete column F (zero‑based index 5) from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteColumn(5);

        // Save the result as a PDF document
        string outputFile = "output.pdf";
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
