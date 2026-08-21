// Title: Save a Modified Excel Workbook in Its Original Format with Aspose.Cells (C#)
// Description: Load an existing workbook, change cell values, detect the source file's format via FileFormatUtil, and save the updated workbook to a new file while preserving all original styles, borders, and shapes.
// Keywords: Aspose.Cells save original format | Workbook.Save same format | FileFormatUtil C# | preserve Excel formatting Aspose | copy workbook with changes
// Common Searches: Aspose.Cells save workbook in original file type | keep original Excel formatting after edit .NET | detect source workbook format Aspose.Cells | save modified workbook without losing styles
// Developer Intent: Export a changed workbook to a new file without altering its original file type or visual layout.
// Use Cases: Update a legacy .xls template and save the result as .xls for older applications. | Apply batch edits to a .xlsx report and generate a copy in the same .xlsx format for downstream workflows. | Create versioned snapshots of a spreadsheet after modifications while retaining all cell styles, charts, and shapes.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, modifies several cells, and saves the workbook using the same format as the source. | Show how to retrieve the original SaveFormat of a loaded workbook and pass it to Workbook.Save to keep all formatting intact. | Explain best practices for preserving cell styles, borders, and embedded objects when saving a modified workbook with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an existing workbook, change cell values, detect the source file's format via FileFormatUtil, and save the updated workbook to a new file while preserving all original styles, borders, and shapes.
class SaveModifiedWorkbook
{
    static void Main()
    {
        // Load the original workbook from disk
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Perform any modifications (example: change a cell value)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["B2"].PutValue("Modified");

        // Get the original file format of the loaded workbook
        SaveFormat originalFormat = FileFormatUtil.FileFormatToSaveFormat(workbook.FileFormat);

        // Save the modified workbook to a new file using the same format
        string destPath = "output_modified.xlsx";
        workbook.Save(destPath, originalFormat);
    }
}
