using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExportCsv
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (source data)
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Food");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Travel");
            dataSheet.Cells["B3"].PutValue(300);
            dataSheet.Cells["A4"].PutValue("Food");
            dataSheet.Cells["B4"].PutValue(80);
            dataSheet.Cells["A5"].PutValue("Travel");
            dataSheet.Cells["B5"].PutValue(150);

            // Add a worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure fields: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table so data is populated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Get the data body range of the pivot table (the area that contains the values)
            CellArea dataBody = pivotTable.DataBodyRange;

            // Determine the range dimensions
            int firstRow = dataBody.StartRow;
            int firstColumn = dataBody.StartColumn;
            int totalRows = dataBody.EndRow - dataBody.StartRow + 1;
            int totalColumns = dataBody.EndColumn - dataBody.StartColumn + 1;

            // Export the range to a DataTable as strings (including column names)
            DataTable dt = pivotSheet.Cells.ExportDataTableAsString(
                firstRow,
                firstColumn,
                totalRows,
                totalColumns,
                true);

            // Write the DataTable to a CSV file (comma‑separated, values quoted)
            using (StreamWriter writer = new StreamWriter("PivotData.csv"))
            {
                foreach (DataRow row in dt.Rows)
                {
                    // Escape double quotes by doubling them and wrap each field in quotes
                    string line = string.Join(",",
                        row.ItemArray.Select(item =>
                            $"\"{item.ToString().Replace("\"", "\"\"")}\""));
                    writer.WriteLine(line);
                }
            }

            // Optionally save the workbook for verification
            workbook.Save("PivotExportDemo.xlsx");
        }
    }
}