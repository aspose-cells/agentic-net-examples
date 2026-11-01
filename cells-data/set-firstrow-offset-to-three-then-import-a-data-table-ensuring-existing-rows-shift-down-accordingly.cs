// Title: Insert three rows at the top of a worksheet and import a DataTable with headers starting at row 4 using Aspose.Cells for .NET
// AI Prompts: Add three blank rows at the beginning of the first worksheet, then write a DataTable (including column names) beginning at cell A4. | Create a reusable helper method that accepts any DataTable, a start row, a start column, and a flag for headers, inserts the required rows, and populates the worksheet accordingly.
// Common Searches: Aspose.Cells how to add rows at the top of an Excel sheet before importing data | C# import DataTable into worksheet starting at row 4 with Aspose.Cells | shift existing rows down in Aspose.Cells .NET before writing a DataTable | Aspose.Cells insert rows and preserve existing data then add DataTable headers
// Tags: Aspose.Cells insert top rows .NET | Aspose.Cells DataTable import with column headers | worksheet row offset before loading DataTable | C# shift worksheet rows down using Aspose.Cells | import DataTable to specific cell range Aspose.Cells

using Aspose.Cells;
using System;
using System.Data;
using System.IO;

// The example loads (or creates) an Excel workbook, inserts three rows at the top of the first worksheet to shift existing content downward, builds a DataTable with Name and Age columns, and imports the table—including column headers—starting at row 4 before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a blank workbook as a placeholder.
            if (!File.Exists(inputPath))
            {
                var placeholder = new Workbook();
                placeholder.Save(inputPath);
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Insert three rows at the top to shift existing rows down.
            sheet.Cells.InsertRows(0, 3);

            // Create and populate a DataTable.
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Age", typeof(int));
            dt.Rows.Add("Alice", 30);
            dt.Rows.Add("Bob", 25);
            dt.Rows.Add("Charlie", 28);

            // Import the DataTable manually (including headers) starting at row index 3, column 0.
            ImportDataTable(sheet, dt, includeColumnNames: true, startRow: 3, startColumn: 0);

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    private static void ImportDataTable(Worksheet sheet, DataTable table, bool includeColumnNames, int startRow, int startColumn)
    {
        int rowIndex = startRow;
        int colIndex = startColumn;

        // Write column headers if required.
        if (includeColumnNames)
        {
            foreach (DataColumn col in table.Columns)
            {
                sheet.Cells[rowIndex, colIndex++].PutValue(col.ColumnName);
            }
            rowIndex++; // Move to the next row for data.
            colIndex = startColumn;
        }

        // Write data rows.
        foreach (DataRow row in table.Rows)
        {
            foreach (object item in row.ItemArray)
            {
                sheet.Cells[rowIndex, colIndex++].PutValue(item);
            }
            rowIndex++;
            colIndex = startColumn;
        }
    }
}
