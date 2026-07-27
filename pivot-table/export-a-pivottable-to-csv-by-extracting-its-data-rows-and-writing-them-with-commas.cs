using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ExportPivotToCsv
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Transport");
        dataSheet.Cells["B3"].PutValue(50);
        dataSheet.Cells["A4"].PutValue("Food");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("Utilities");
        dataSheet.Cells["B5"].PutValue(200);

        // Add a worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range, destination cell, name)
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B5", "A1", "MyPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Populate the pivot table
        pivotTable.CalculateData();

        // Determine the data body range of the pivot table
        CellArea dataBody = pivotTable.DataBodyRange;
        int startRow = dataBody.StartRow;
        int startCol = dataBody.StartColumn;
        int totalRows = dataBody.EndRow - dataBody.StartRow + 1;
        int totalCols = dataBody.EndColumn - dataBody.StartColumn + 1;

        // Export the pivot data to a DataTable (as strings, include column names)
        DataTable dt = pivotSheet.Cells.ExportDataTableAsString(startRow, startCol, totalRows, totalCols, true);

        // Build CSV content from the DataTable
        StringBuilder csvBuilder = new StringBuilder();

        // Write header row
        foreach (DataColumn col in dt.Columns)
        {
            csvBuilder.Append(col.ColumnName);
            csvBuilder.Append(',');
        }
        csvBuilder.Length--; // Remove trailing comma
        csvBuilder.AppendLine();

        // Write data rows
        foreach (DataRow row in dt.Rows)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                csvBuilder.Append(row[i].ToString());
                csvBuilder.Append(',');
            }
            csvBuilder.Length--; // Remove trailing comma
            csvBuilder.AppendLine();
        }

        // Save CSV file
        string csvPath = "PivotData.csv";
        File.WriteAllText(csvPath, csvBuilder.ToString());

        // Optionally save the workbook for reference
        workbook.Save("PivotWorkbook.xlsx");
    }
}