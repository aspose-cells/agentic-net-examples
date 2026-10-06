// Title: Insert rows at the end of a worksheet and auto‑fit all row heights using Aspose.Cells for .NET (C#)
// AI Prompts: Load or create a workbook, locate the final populated row, add two new rows, write sample data, then invoke sheet.AutoFitRows(startRow, totalRows) to resize every row. | With Aspose.Cells in C#, programmatically extend the data range by inserting rows and automatically calibrate the height of all rows in the sheet.
// Common Searches: c# aspocells insert rows after last data row and autofit row heights | how to use AutoFitRows on a dynamic range in Aspose.Cells .NET | adjust all row heights after adding new rows with Aspose.Cells for .NET | Aspose.Cells C# example for inserting rows and applying AutoFitRows | auto fit rows in newly inserted rows in Excel using Aspose.Cells
// Tags: InsertRows method Aspose.Cells C# | AutoFitRows overload Aspose.Cells | adjust row heights .xlsx Aspose.Cells | dynamic worksheet row insertion Aspose.Cells | maintain uniform row height Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, determines the last used row, inserts two rows, populates them with sample values, and calls sheet.AutoFitRows(0, sheet.Cells.MaxDataRow + 1) to automatically adjust the height of every row before saving the file.
class AutoFitRowsExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the row after the last used row.
            int insertRowIndex = sheet.Cells.MaxDataRow + 1;

            // Insert two new rows at the determined position.
            sheet.Cells.InsertRows(insertRowIndex, 2);

            // Populate the newly inserted rows with sample data.
            sheet.Cells[insertRowIndex, 0].PutValue("New Data Row 1");
            sheet.Cells[insertRowIndex, 1].PutValue(123);
            sheet.Cells[insertRowIndex + 1, 0].PutValue("New Data Row 2");
            sheet.Cells[insertRowIndex + 1, 1].PutValue(456);

            // Apply AutoFitRows to all rows that contain data.
            int startRow = 0;
            int totalRows = sheet.Cells.MaxDataRow + 1; // number of rows to autofit

            // AutoFitRows overload with start row and total rows.
            sheet.AutoFitRows(startRow, totalRows);

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
