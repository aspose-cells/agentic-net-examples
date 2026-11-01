// Title: Insert two rows at row 51 in an Excel workbook with Aspose.Cells for .NET and save the updated file as PDF
// AI Prompts: Use Aspose.Cells for .NET to add two rows after row 50 in an existing .xlsx file while keeping the original formatting, then export the workbook to a PDF. | Write C# code that loads an Excel workbook, inserts rows at index 50 preserving default styles, and saves the modified document as a PDF using Aspose.Cells.
// Common Searches: how to add rows at a specific index in an Excel file using Aspose.Cells C# | Aspose.Cells insert rows preserve formatting then convert to PDF | C# example for inserting rows in .xlsx and exporting to PDF with Aspose.Cells | insert multiple rows at row 51 using Aspose.Cells for .NET | convert modified Excel workbook to PDF after inserting rows Aspose.Cells
// Tags: insert rows Aspose.Cells C# | preserve formatting when inserting rows Aspose.Cells | export Excel to PDF Aspose.Cells .NET | modify workbook then save as PDF Aspose.Cells | row insertion index 50 Aspose.Cells

using System;
using Aspose.Cells;

// Loads an existing .xlsx workbook, inserts two rows at zero‑based index 50 (row 51) while copying surrounding formatting, and saves the result as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Insert two rows at index 50 (zero‑based index, i.e., after row 49)
        // This operation copies the default formatting from the surrounding rows.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.InsertRows(50, 2);

        // Save the modified workbook as a PDF (replace with desired output path)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
