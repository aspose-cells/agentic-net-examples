using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

class QueryTableReadWriteDemo
{
    static void Main()
    {
        // Load the existing workbook that contains a query table
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // lifecycle: load

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Ensure there is at least one query table
        if (worksheet.QueryTables.Count == 0)
        {
            Console.WriteLine("No query tables found in the worksheet.");
            return;
        }

        // Get the first query table
        QueryTable queryTable = worksheet.QueryTables[0];

        // Export the result range of the query table to a DataTable
        // ExportDataTable() rule is used here
        DataTable dataTable = queryTable.ResultRange.ExportDataTable();

        // ----- Perform tabular operations on the DataTable -----
        // Example: add a new row with sample data
        DataRow newRow = dataTable.NewRow();
        // Assuming the query table has at least two columns; adjust as needed
        if (dataTable.Columns.Count >= 2)
        {
            newRow[0] = "NewItem";
            newRow[1] = 12345;
        }
        dataTable.Rows.Add(newRow);

        // ----- Write the modified DataTable back to the worksheet -----
        // Determine the start position of the original result range
        int startRow = queryTable.ResultRange.FirstRow;
        int startColumn = queryTable.ResultRange.FirstColumn;

        // Prepare import options (show field names and insert rows if needed)
        ImportTableOptions importOptions = new ImportTableOptions
        {
            IsFieldNameShown = true,
            InsertRows = true
        };

        // ImportData(DataTable, int, int, ImportTableOptions) rule is used here
        cells.ImportData(dataTable, startRow, startColumn, importOptions);

        // Save the workbook with the updated data
        string outputPath = "output.xlsx";
        workbook.Save(outputPath); // lifecycle: save

        Console.WriteLine($"Query table data read, modified, and saved to '{outputPath}'.");
    }
}