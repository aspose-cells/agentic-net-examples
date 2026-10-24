// Title: How to set Excel table header cells using ListObject.PutCellValue with offsets in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set table header using ListObject.PutCellValue row offset | populate ListObject header row in Excel workbook with Aspose.Cells .NET | use PutCellValue to write header values in an Aspose.Cells table | Aspose.Cells .NET example adding header to ListObject table programmatically | write Excel table column headers with ListObject.PutCellValue offsets
// Tags: ListObject.PutCellValue header assignment | Aspose.Cells table header creation | C# Excel table cell offset writing | Aspose.Cells .xlsx table population | Set ListObject header values programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds a ListObject table spanning five rows and three columns, and uses ListObject.PutCellValue with row and column offsets (0,0), (0,1), and (0,2) to assign values to the three header cells before saving the file as TableWithHeader.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range for the ListObject (table)
            // Table starts at cell A1 (row 0, column 0) and spans 5 rows and 3 columns (including header)
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = 5;      // includes header row
            int totalColumns = 3;

            // Add a ListObject to the worksheet (the table)
            int listObjectIndex = worksheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows,
                firstColumn + totalColumns,
                true);

            ListObject table = worksheet.ListObjects[listObjectIndex];

            // Populate header cells using PutCellValue (row offset, column offset, value)
            table.PutCellValue(0, 0, "Header1");
            table.PutCellValue(0, 1, "Header2");
            table.PutCellValue(0, 2, "Header3");

            // Save the workbook
            string outputPath = "TableWithHeader.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
