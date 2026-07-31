// Title: C# Aspose.Cells – Create a Growing Named Range for an Excel Table
// Description: The sample builds a workbook, adds a header and rows, creates a ListObject on A1:B4, and defines a named range with an OFFSET‑COUNTA formula that tracks column A. After inserting a new row, the table is resized, formulas are recalculated, and the updated range address is printed before saving the file.
// Keywords: Aspose.Cells | C# dynamic named range | OFFSET COUNTA | Excel table resizing | ListObject expand | auto‑growing range | named range formula | add row programmatically | Aspose.Cells workbook | Excel automation C#
// Common Searches: Aspose.Cells dynamic range example | C# expand ListObject after adding rows | OFFSET COUNTA named range Aspose.Cells | how to auto‑grow named range in Excel with code | programmatically resize Excel table using Aspose | retrieve updated named range address C#
// Developer Intent: Generate a table whose named range automatically grows when rows are added.
// Use Cases: Set up a ListObject and a self‑adjusting range for column‑based data. | Programmatically insert additional rows and automatically extend the table size. | Maintain accurate range references for formulas or data validation after each insertion. | Export the workbook with the expanded range for downstream processing.
// AI Prompts: Write C# Aspose.Cells code that creates an Excel table and a dynamic named range using OFFSET and COUNTA. | Explain the role of OFFSET and COUNTA in making a named range responsive to new rows. | Show how to add a row to a ListObject and resize it so the associated named range updates. | Provide a step‑by‑step guide to retrieve the address of a dynamic named range after modifying the table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The sample builds a workbook, adds a header and rows, creates a ListObject on A1:B4, and defines a named range with an OFFSET‑COUNTA formula that tracks column A. After inserting a new row, the table is resized, formulas are recalculated, and the updated range address is printed before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Populate initial data -----
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Value");

            // Sample data rows (A2:B4)
            for (int i = 2; i <= 4; i++)
            {
                sheet.Cells[i - 1, 0].PutValue(i - 1);          // ID
                sheet.Cells[i - 1, 1].PutValue(i * 10);        // Value
            }

            // ----- Create a ListObject (table) covering the data -----
            // Table range: A1:B4
            int tableIdx = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIdx];
            table.DisplayName = "MyTable";

            // ----- Define a dynamic named range that expands with column A -----
            // The formula uses OFFSET and COUNTA to count non‑empty cells in column A.
            // It starts from A2 (first data cell) and expands downwards.
            int nameIdx = workbook.Worksheets.Names.Add("DynamicIDs");
            Name dynamicName = workbook.Worksheets.Names[nameIdx];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$2,0,0,COUNTA({sheet.Name}!$A:$A)-1,1)";

            // ----- Add a new row of data below the current table -----
            // Determine the row index after the current table
            int dataRangeFirstRow = table.DataRange.FirstRow;
            int dataRangeRowCount = table.DataRange.RowCount;
            int newDataRow = dataRangeFirstRow + dataRangeRowCount; // row index after the table

            sheet.Cells[newDataRow, 0].PutValue(5);   // New ID
            sheet.Cells[newDataRow, 1].PutValue(50); // New Value

            // Resize the table to include the newly added row
            table.Resize(table.StartRow, table.StartColumn, newDataRow, table.EndColumn, true);

            // Recalculate formulas (if any)
            workbook.CalculateFormula();

            // Retrieve and display the address of the expanded dynamic range
            AsposeRange expandedRange = dynamicName.GetRange();
            Console.WriteLine($"Dynamic range address after adding row: {expandedRange.Address}");

            // Save the workbook
            string outputPath = "DynamicTableNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
