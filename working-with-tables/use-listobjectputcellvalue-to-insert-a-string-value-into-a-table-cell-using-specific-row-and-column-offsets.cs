// Title: How to replace a table cell value with a string using ListObject.PutCellValue in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set value in ListObject table cell by row and column offset | Replace specific cell in an Aspose.Cells ListObject with a string | Use ListObject.PutCellValue to update table data in a .NET workbook | Change value of cell B3 in an Aspose.Cells table programmatically | Insert string into Aspose.Cells ListObject using row and column indices
// Tags: ListObject.PutCellValue C# | update Aspose.Cells table cell | modify ListObject cell value | Aspose.Cells table row column offset | replace cell value in Excel workbook .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a range A1:C4 as a ListObject table, and then uses ListObject.PutCellValue with row and column offsets (1,1) to replace the existing value in cell B3 with the string "Robert" before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data that will become the table (range A1:C4)
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(90);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Define the range for the ListObject (table) – rows 0‑3, columns 0‑2 (A1:C4)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 4;
            int totalColumns = 3;

            // Add the ListObject to the worksheet (hasHeaders = true)
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            // Obtain a reference to the newly added table
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Replace the value in the second data row, second column (cell B3) with "Robert"
            table.PutCellValue(1, 1, "Robert");

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
