// Title: Auto‑fit all columns of an Aspose.Cells ListObject (table) to the longest cell content using C#
// AI Prompts: Generate C# code that creates a worksheet, adds a ListObject, and calls AutoFitColumn on every column of the table's data range. | Show how to iterate through the columns of a ListObject in Aspose.Cells and automatically resize each column to the widest cell value. | Provide a complete example that populates sample data, defines a table, and saves the workbook after auto‑sizing the table columns.
// Common Searches: Aspose.Cells C# auto fit ListObject column width to content | How to resize columns of a table based on longest text in Aspose.Cells .NET | C# example for auto‑sizing columns of an Aspose.Cells table | AutoFitColumn for table data range Aspose.Cells tutorial | Adjust column widths of a ListObject to fit cell values using Aspose.Cells
// Tags: Aspose.Cells ListObject column auto‑fit | C# resize table columns to content | ListObject column width adjustment | auto‑fit columns using data range | Aspose.Cells column width optimization

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;

// The program creates a new workbook, adds a three‑column ListObject with sample headers and data, iterates over the table's data range columns, calls AutoFitColumn on each column to match the longest cell content, and saves the result as AutoFitTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["A2"].PutValue("Short");
            sheet.Cells["B2"].PutValue("A much longer text that should expand column width");
            sheet.Cells["C2"].PutValue("Mid");

            // Define the range that will become a table (ListObject)
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // zero‑based index for column A
            int totalRows = 2;  // header + one data row
            int totalCols = 3;  // three columns

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstCol,
                firstRow + totalRows - 1, firstCol + totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Auto‑fit all columns of the table to match the longest cell content
            // Use DataRange (includes data rows) to determine column boundaries
            Aspose.Cells.Range dataRange = table.DataRange;
            int startCol = dataRange.FirstColumn;
            int endCol = dataRange.FirstColumn + dataRange.ColumnCount - 1;
            for (int col = startCol; col <= endCol; col++)
            {
                sheet.AutoFitColumn(col);
            }

            // Save the workbook
            workbook.Save("AutoFitTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
