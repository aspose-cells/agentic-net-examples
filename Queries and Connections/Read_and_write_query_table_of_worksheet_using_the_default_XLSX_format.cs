using System;
using Aspose.Cells;

class QueryTableReadWriteDemo
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data (simulating external data source)
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("John");
        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Mary");

        // The QueryTables collection is read‑only; we can only read existing tables.
        // Check if any query tables are present.
        if (worksheet.QueryTables.Count > 0)
        {
            // Read properties of the first query table
            QueryTable qt = worksheet.QueryTables[0];
            Console.WriteLine("Query Table Name: " + qt.Name);
            Console.WriteLine("Result Range: " + qt.ResultRange.Address);
            Console.WriteLine("PreserveFormatting (before): " + qt.PreserveFormatting);
            Console.WriteLine("AdjustColumnWidth (before): " + qt.AdjustColumnWidth);

            // Modify (write) some properties
            qt.PreserveFormatting = true;
            qt.AdjustColumnWidth = false;

            Console.WriteLine("PreserveFormatting (after): " + qt.PreserveFormatting);
            Console.WriteLine("AdjustColumnWidth (after): " + qt.AdjustColumnWidth);
        }
        else
        {
            Console.WriteLine("No query tables found in the worksheet.");
        }

        // Save the workbook in the default XLSX format (save rule)
        workbook.Save("QueryTableReadWriteDemo.xlsx");
    }
}