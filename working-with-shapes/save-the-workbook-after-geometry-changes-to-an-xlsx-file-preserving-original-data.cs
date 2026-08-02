// Title: Insert a Row in an Existing XLSX and Save with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX file with Aspose.Cells, inserts a new row at a specified index using InsertRange with ShiftType.Down, optionally writes values to the new cells, and saves the workbook as a new XLSX while keeping all original data intact.
// Keywords: Aspose.Cells C# | insert row Excel | InsertRange ShiftType.Down | save workbook XLSX | preserve existing data | worksheet geometry | .NET Excel automation | load workbook Aspose | Excel file modification
// Common Searches: Aspose.Cells insert row and keep data | C# save Excel after geometry changes | How to add a row with Aspose.Cells .NET | InsertRange example Aspose.Cells | Preserve data when inserting rows in XLSX
// Developer Intent: Add a new row to an existing Excel file and write the updated workbook without losing any prior content.
// Use Cases: Add a header row to a template before distribution. | Create blank rows for user input while retaining existing records. | Insert a summary or total row at a specific position in a report.
// AI Prompts: Generate C# code that loads an XLSX with Aspose.Cells, inserts multiple rows at index 5 using ShiftType.Down, fills the new cells, and saves the file preserving all data. | Show an example of using Aspose.Cells for .NET to modify worksheet geometry (insert row) and then export the workbook as XLSX.

using System;
using Aspose.Cells;

namespace AsposeCellsGeometrySaveDemo
{
    // Loads an XLSX file with Aspose.Cells, inserts a new row at a specified index using InsertRange with ShiftType.Down, optionally writes values to the new cells, and saves the workbook as a new XLSX while keeping all original data intact.
    public class Program
    {
        public static void Main()
        {
            // Path to the existing workbook
            string sourcePath = "input.xlsx";

            // Load the workbook from the file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a cell area where we want to insert a new row (row index 3, columns 0‑1)
            // This will shift existing rows down, preserving data below the insertion point
            CellArea insertArea = CellArea.CreateCellArea(3, 0, 3, 1);
            sheet.Cells.InsertRange(insertArea, 1, ShiftType.Down, true);

            // Optionally, add data to the newly inserted row
            sheet.Cells["A4"].PutValue("New Row");
            sheet.Cells["B4"].PutValue(123);

            // Save the modified workbook to a new XLSX file, preserving all original data
            // Uses Workbook.Save(string, SaveFormat) method
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook geometry updated and saved to output.xlsx");
        }
    }
}
