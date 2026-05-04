using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsTypedColumnsDemo
{
    // Simple POCO representing a product row
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];              // get first sheet
            Cells cells = sheet.Cells;                             // shortcut to cells collection

            // ------------------------------------------------------------
            // 2. Define column headers (strongly typed metadata)
            // ------------------------------------------------------------
            // Header row (row 0)
            cells[0, 0].PutValue("Id");          // integer column
            cells[0, 1].PutValue("Name");        // string column
            cells[0, 2].PutValue("Price");       // double column
            cells[0, 3].PutValue("ReleaseDate"); // DateTime column

            // ------------------------------------------------------------
            // 3. Populate sample data with correct .NET types
            // ------------------------------------------------------------
            var sampleData = new List<Product>
            {
                new Product { Id = 1, Name = "Apple",  Price = 0.99,  ReleaseDate = new DateTime(2023,12,01) },
                new Product { Id = 2, Name = "Banana", Price = 1.49,  ReleaseDate = new DateTime(2023,12,05) },
                new Product { Id = 3, Name = "Cherry", Price = 2.99,  ReleaseDate = new DateTime(2023,12,10) }
            };

            // Import the list into the worksheet starting at row 1 (below headers)
            // Use ImportCustomObjects overload that respects property names
            string[] propertyNames = { "Id", "Name", "Price", "ReleaseDate" };
            cells.ImportCustomObjects(
                sampleData,
                propertyNames,
                false,          // do NOT write property names again (headers already present)
                1,              // first data row
                0,              // first column
                sampleData.Count,
                true,           // insert rows if needed
                "yyyy-MM-dd",   // date format
                true);          // convert strings to numbers when possible

            // ------------------------------------------------------------
            // 4. Export the typed range to a DataTable with type checking
            // ------------------------------------------------------------
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,   // first row contains column names
                CheckMixedValueType = true // enforce consistent column types
            };

            // Export rows 0..(sampleData.Count) and columns 0..3
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, sampleData.Count + 1, 4, exportOptions);

            // Display the inferred .NET types of each column
            Console.WriteLine("Exported DataTable column types:");
            foreach (DataColumn col in dt.Columns)
            {
                Console.WriteLine($"{col.ColumnName} : {col.DataType}");
            }

            // ------------------------------------------------------------
            // 5. Demonstrate type safety by attempting to add a mixed type row
            // ------------------------------------------------------------
            // This row mixes string into the Id column (should become string if CheckMixedValueType were false)
            sheet.Cells[4, 0].PutValue("MixedId");
            sheet.Cells[4, 1].PutValue("Invalid");
            sheet.Cells[4, 2].PutValue(9.99);
            sheet.Cells[4, 3].PutValue(DateTime.Now);

            // Re‑export with the same options; the Id column will now be typed as string because of the mixed value
            DataTable dtMixed = sheet.Cells.ExportDataTable(0, 0, sampleData.Count + 2, 4, exportOptions);
            Console.WriteLine("\nAfter inserting a mixed‑type row:");
            foreach (DataColumn col in dtMixed.Columns)
            {
                Console.WriteLine($"{col.ColumnName} : {col.DataType}");
            }

            // ------------------------------------------------------------
            // 6. Clean the sheet and re‑import the original strongly‑typed data
            // ------------------------------------------------------------
            sheet.Cells.Clear(); // remove all existing content

            // Re‑write headers
            cells[0, 0].PutValue("Id");
            cells[0, 1].PutValue("Name");
            cells[0, 2].PutValue("Price");
            cells[0, 3].PutValue("ReleaseDate");

            // Import only the original strongly‑typed list (no mixed row)
            cells.ImportCustomObjects(
                sampleData,
                propertyNames,
                false,
                1,
                0,
                sampleData.Count,
                true,
                "yyyy-MM-dd",
                true);

            // ------------------------------------------------------------
            // 7. Save the workbook as XLSX
            // ------------------------------------------------------------
            workbook.Save("TypedColumnsDemo.xlsx"); // save workbook
        }
    }
}