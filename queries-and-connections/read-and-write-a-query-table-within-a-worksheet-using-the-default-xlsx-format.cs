using System;
using Aspose.Cells;

namespace AsposeCellsQueryTableDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook (lifecycle rule: create) ----------
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data (this will be the area that could be a query table)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("John");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Mary");

            // ---------- Read the QueryTable collection ----------
            QueryTableCollection queryTables = worksheet.QueryTables;

            if (queryTables.Count > 0)
            {
                // There is at least one query table; read its properties
                QueryTable qt = queryTables[0];
                Console.WriteLine("Query Table Name: " + qt.Name);
                Console.WriteLine("Result Range: " + qt.ResultRange.Address);
                Console.WriteLine("Preserve Formatting (before): " + qt.PreserveFormatting);

                // ---------- Write (modify) a property of the QueryTable ----------
                // Toggle PreserveFormatting to demonstrate a write operation
                qt.PreserveFormatting = !qt.PreserveFormatting;
                Console.WriteLine("Preserve Formatting (after): " + qt.PreserveFormatting);
            }
            else
            {
                // No query tables exist; inform the user
                Console.WriteLine("No query tables found in the worksheet.");
                // In a real scenario, a query table would be created via an external data source.
                // Here we simply demonstrate that the collection is accessible.
            }

            // ---------- Save the workbook (lifecycle rule: save) ----------
            workbook.Save("QueryTableDemo.xlsx");
        }
    }
}