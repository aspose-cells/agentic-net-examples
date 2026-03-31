using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataTableToXlsx
{
    class Program
    {
        static void Main()
        {
            // ---------- 1. Prepare a DataTable with various column types ----------
            DataTable dt = new DataTable();

            // Define columns with explicit data types
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(double));
            dt.Columns.Add("ReleaseDate", typeof(DateTime));

            // Add sample rows
            dt.Rows.Add("Laptop", 10, 1299.99, new DateTime(2023, 1, 15));
            dt.Rows.Add("Smartphone", 25, 799.5, new DateTime(2023, 3, 5));
            dt.Rows.Add("Tablet", 15, 450.75, new DateTime(2022, 11, 20));

            // ---------- 2. Create a new workbook (create rule) ----------
            Workbook workbook = new Workbook(); // uses the default constructor

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ---------- 3. Import the DataTable into the worksheet ----------
            // ImportData will map the DataTable rows to cells starting at row 0, column 0
            // ImportTableOptions can be customized; using default options here
            cells.ImportData(dt, 0, 0, new ImportTableOptions());

            // ---------- 4. Preserve column formatting based on DataTable column types ----------
            // Apply number formats for numeric and date columns
            for (int col = 0; col < dt.Columns.Count; col++)
            {
                // Create a style object for the column
                Style colStyle = workbook.CreateStyle();

                // Determine the appropriate number format
                Type dataType = dt.Columns[col].DataType;
                if (dataType == typeof(int) || dataType == typeof(double) || dataType == typeof(decimal))
                {
                    // Two decimal places for monetary values, integer format for whole numbers
                    colStyle.Custom = dataType == typeof(int) ? "0" : "#,##0.00";
                }
                else if (dataType == typeof(DateTime))
                {
                    // Standard date format
                    colStyle.Custom = "mm-dd-yyyy";
                }
                // Apply the style to the entire column (including header)
                cells.ApplyColumnStyle(col, colStyle, new StyleFlag() { NumberFormat = true });
            }

            // ---------- 5. Save the workbook to an XLSX file (save rule) ----------
            workbook.Save("DataTableExport.xlsx"); // saves in XLSX format by default
        }
    }
}