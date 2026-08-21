// Title: Insert a Row and Save an XLSX Workbook with Aspose.Cells for .NET
// Description: Load an existing XLSX file using Aspose.Cells, insert a new row at a specified index with InsertRange (ShiftType.Down), and save the workbook back to XLSX while preserving all data, formulas, formatting, and merged cells.
// Keywords: Aspose.Cells insert row C# | InsertRange ShiftType.Down | save workbook after row insertion | preserve formulas Aspose.Cells | C# Excel row insertion example | Aspose.Cells .NET save XLSX | Excel geometry changes Aspose
// Common Searches: How to add a row to an existing Excel file with Aspose.Cells .NET | Saving an Excel workbook after inserting rows without losing data | Aspose.Cells InsertRange example for shifting rows down | C# code to insert a row and keep formatting in XLSX
// Developer Intent: Add a new row to an existing XLSX workbook and save the file while keeping all original content intact.
// Use Cases: Insert a header row at the top of a report before exporting to XLSX. | Add a blank separator row within a data table to improve readability without breaking formulas. | Shift schedule rows down to accommodate new entries while preserving cell styles and merged regions.
// AI Prompts: Generate C# code that inserts multiple rows at a given index using Aspose.Cells and saves the workbook preserving formulas and formatting. | Explain how InsertRange handles merged cells and data validation when inserting a row with Aspose.Cells for .NET. | Show how to insert a row and save the workbook to a MemoryStream instead of a physical file.

using System;
using Aspose.Cells;

// Load an existing XLSX file using Aspose.Cells, insert a new row at a specified index with InsertRange (ShiftType.Down), and save the workbook back to XLSX while preserving all data, formulas, formatting, and merged cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Example geometry change: insert a new row at index 3 (fourth row)
        // Define the range that represents the entire row to be shifted down
        CellArea insertArea = CellArea.CreateCellArea(3, 0, 3, worksheet.Cells.MaxColumn);
        worksheet.Cells.InsertRange(insertArea, 3, ShiftType.Down, true);

        // Save the modified workbook to XLSX format, preserving all original data
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
