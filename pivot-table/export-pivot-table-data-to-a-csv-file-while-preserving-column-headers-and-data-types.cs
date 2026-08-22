// Title: Export a PivotTable to CSV with column headers and native data types using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to export a PivotTable report to a CSV file, keeping the original column names and preserving numeric and date types. | Show how to configure ExportTableOptions (ExportColumnName = true, ExportAsString = false) to extract a PivotTable into a DataTable and then write it as a CSV with proper type handling.
// Common Searches: c# aspnet export aspose.cells pivot table to csv preserving data types | how to keep column headers when exporting Aspose.Cells pivot report to CSV | ExportTableOptions ExportAsString false example in Aspose.Cells | write DataTable to CSV with invariant culture in C# Aspose.Cells pivot | save pivot table data as CSV file using Aspose.Cells .NET
// Tags: Aspose.Cells export pivot table to CSV | ExportTableOptions preserve data types | PivotTable TableRange1 to DataTable | C# write DataTable as CSV with invariant culture | include column headers in Aspose.Cells CSV export

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, builds a pivot table, uses ExportTableOptions (ExportColumnName=true, ExportAsString=false) to export the pivot report (TableRange1) to a DataTable, then writes the DataTable to a CSV file while preserving column headers and native numeric/date types, and saves the workbook for verification.
class ExportPivotToCsv
{
    static void Main()
    {
        // Create a workbook and add source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["A3"].PutValue("Transport");
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["A4"].PutValue("Food");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("Transport");
        dataSheet.Cells["B5"].PutValue(70);

        // Add a worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Define the source range for the pivot table
        string sourceRange = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";

        // Create the pivot table
        int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A1", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Set export options: include column headers and keep original data types
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            ExportColumnName = true,
            ExportAsString = false
        };

        // Export the pivot table report (TableRange1) to a DataTable
        CellArea reportArea = pivotTable.TableRange1;
        DataTable dataTable = workbook.Worksheets[pivotSheet.Index].Cells.ExportDataTable(
            reportArea.StartRow,
            reportArea.StartColumn,
            reportArea.EndRow - reportArea.StartRow + 1,
            reportArea.EndColumn - reportArea.StartColumn + 1,
            exportOptions);

        // Write the DataTable to a CSV file, preserving data types
        string csvFilePath = "PivotExport.csv";
        using (StreamWriter writer = new StreamWriter(csvFilePath))
        {
            // Write column headers
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                writer.Write(dataTable.Columns[col].ColumnName);
                if (col < dataTable.Columns.Count - 1) writer.Write(",");
            }
            writer.WriteLine();

            // Write each row
            foreach (DataRow row in dataTable.Rows)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    object value = row[col];
                    if (value is string str)
                    {
                        // Escape commas and quotes in string values
                        if (str.Contains(",") || str.Contains("\""))
                        {
                            str = $"\"{str.Replace("\"", "\"\"")}\"";
                        }
                        writer.Write(str);
                    }
                    else
                    {
                        // Write numeric/date values using invariant culture
                        writer.Write(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));
                    }

                    if (col < dataTable.Columns.Count - 1) writer.Write(",");
                }
                writer.WriteLine();
            }
        }

        // Save the workbook (optional, for verification)
        workbook.Save("PivotWorkbook.xlsx");
    }
}
