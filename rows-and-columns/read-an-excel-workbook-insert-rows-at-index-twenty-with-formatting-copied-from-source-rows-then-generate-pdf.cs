// Title: Insert rows at a specific index in an Excel worksheet, copy formatting from shifted rows, and export the workbook to PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Insert one row at zero‑based index 20 in an existing Excel file, copy the formatting of the rows that are shifted down, and save the workbook as a PDF with Aspose.Cells for .NET. | Add multiple consecutive rows at a given position in a worksheet, preserve the original cell styles of the shifted rows, and generate a PDF output using Aspose.Cells.
// Common Searches: Aspose.Cells C# insert row at row 21 keep formatting | how to add rows to an Excel file and export to PDF with Aspose.Cells .NET | copy cell style when inserting rows using Aspose.Cells for .NET | insert rows at specific index and save workbook as PDF in C# | preserve formatting during row insertion Aspose.Cells
// Tags: insert rows with formatting Aspose.Cells | worksheet.Cells.InsertRows C# | export workbook to PDF Aspose.Cells | zero‑based row index insertion Excel .NET | preserve cell styles during row insertion | Aspose.Cells row insertion PDF conversion

using Aspose.Cells;
using System;

// The program loads 'input.xlsx', inserts one row at zero‑based index 20 while copying the formatting of the rows that are shifted down, and saves the modified workbook as 'output.pdf' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the row index where new rows will be inserted (zero‑based, so 20 = row 21 in Excel)
        int insertRowIndex = 20;

        // Number of rows to insert
        int rowsToInsert = 1;   // change as required

        // Insert rows and copy formatting from the source rows that are being shifted down
        worksheet.Cells.InsertRows(insertRowIndex, rowsToInsert, true);

        // Save the modified workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
