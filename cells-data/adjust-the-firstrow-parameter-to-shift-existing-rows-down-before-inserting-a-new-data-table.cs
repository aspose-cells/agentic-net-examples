// Title: Insert blank rows and create a data table at a specific row with Aspose.Cells for .NET (C#)
// AI Prompts: Insert three empty rows at row index 5 in an existing workbook, then write a header row and two data rows starting at that position using Aspose.Cells for .NET. | Adjust the firstRow variable to shift existing rows down before adding a new table, and save the workbook with Aspose.Cells in C#. | Programmatically add a new data table after inserting rows in a worksheet and persist the changes to a new Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# insert rows at specific index and then add table | How to shift rows down before creating a table in Excel with Aspose.Cells | C# code to add blank rows and populate a data table using Aspose.Cells | Insert rows and write header row in Aspose.Cells workbook example
// Tags: insert rows Aspose.Cells C# | shift worksheet rows down Aspose.Cells | add data table to worksheet Aspose.Cells | save modified workbook Aspose.Cells | firstRow parameter usage Aspose.Cells

using Aspose.Cells;

// The example loads an existing workbook, inserts three empty rows at zero‑based row index 5 on the first worksheet, writes a header row and two data rows starting at the adjusted position, and saves the result as a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (load rule)
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Define the row where the new data table should start
        int firstRow = 5;          // original position (0‑based index)
        int rowsToInsert = 3;      // number of rows the new table will occupy

        // Shift existing rows down by inserting empty rows at firstRow (firstRow parameter adjusted)
        sheet.Cells.InsertRows(firstRow, rowsToInsert);

        // Insert the new data table starting at the adjusted firstRow
        // Header row
        sheet.Cells[firstRow, 0].PutValue("Header1");
        sheet.Cells[firstRow, 1].PutValue("Header2");

        // Data rows
        sheet.Cells[firstRow + 1, 0].PutValue("Data1");
        sheet.Cells[firstRow + 1, 1].PutValue("Data2");

        sheet.Cells[firstRow + 2, 0].PutValue("Data3");
        sheet.Cells[firstRow + 2, 1].PutValue("Data4");

        // Save the modified workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
