// Title: Aspose.Cells .NET: Import a DataTable and Auto‑Create an Excel Table with Headers
// Description: Demonstrates how to build a DataTable, import it into the first worksheet using ImportTableOptions.IsFieldNameShown, convert the range to a ListObject (Excel table) that inherits column names as headers, and save the workbook as .xlsx.
// Keywords: Aspose.Cells | C# | DataTable to Excel | ImportData | ImportTableOptions | ListObject | Excel table from DataTable | auto header mapping | worksheet table creation | export DataTable to .xlsx
// Common Searches: Aspose.Cells import DataTable as Excel table | C# create ListObject from DataTable | How to map DataTable columns to Excel table headers | ImportTableOptions.IsFieldNameShown example | Convert DataTable to worksheet table Aspose
// Developer Intent: Generate an Excel worksheet table directly from a DataTable, using the DataTable's column names as the table's header row.
// Use Cases: Export database query results to a ready‑to‑filter Excel table for business reports. | Programmatically build sales or inventory worksheets where column definitions are defined in code. | Provide downstream analysts with a pre‑formatted .xlsx file that includes proper table headers for pivot tables and charts.
// AI Prompts: Write C# code that takes a DataTable, imports it into an Aspose.Cells worksheet, and creates a ListObject with column names as headers. | Show how to enable ImportTableOptions.IsFieldNameShown and add a ListObject so the DataTable schema becomes an Excel table. | Explain the steps to calculate the data range, add the table, and save the workbook after converting a DataTable to an Excel table.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsDataTableToTableDemo
{
    // Demonstrates how to build a DataTable, import it into the first worksheet using ImportTableOptions.IsFieldNameShown, convert the range to a ListObject (Excel table) that inherits column names as headers, and save the workbook as .xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Prepare a DataTable ----------
                DataTable dt = new DataTable("SampleData");
                dt.Columns.Add("Product", typeof(string));
                dt.Columns.Add("Quantity", typeof(int));
                dt.Columns.Add("Price", typeof(double));

                dt.Rows.Add("Apple", 10, 0.5);
                dt.Rows.Add("Banana", 20, 0.3);
                dt.Rows.Add("Cherry", 15, 0.8);

                // ---------- 2. Create a new workbook ----------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];              // default first sheet
                Cells cells = sheet.Cells;

                // ---------- 3. Import the DataTable into the worksheet ----------
                // ImportTableOptions.IsFieldNameShown = true writes column names as the first row.
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true
                };
                cells.ImportData(dt, 0, 0, importOptions);             // import

                // ---------- 4. Convert the imported range into an Excel table ----------
                int totalRows = dt.Rows.Count + 1; // +1 for header row
                int totalCols = dt.Columns.Count;
                // The last parameter 'true' indicates that the first row contains headers.
                sheet.ListObjects.Add(0, 0, totalRows, totalCols, true); // create table

                // ---------- 5. Save the workbook ----------
                string outputPath = "DataTableToExcelTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
