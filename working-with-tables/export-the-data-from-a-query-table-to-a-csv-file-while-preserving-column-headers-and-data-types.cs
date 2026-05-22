using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ExportQueryTableToCsv
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the worksheet with sample data (including headers)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Amount");
            cells["D1"].PutValue("Date");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(123.45);
            cells["D2"].PutValue(new DateTime(2023, 1, 15));

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(678.90);
            cells["D3"].PutValue(new DateTime(2023, 2, 20));

            // Define a ListObject (query table) that covers the data range
            int totalRows = 3;   // header + 2 data rows
            int totalCols = 4;   // four columns
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, totalRows, totalCols, true);
            ListObject table = worksheet.ListObjects[listObjectIndex];
            table.DisplayName = "SampleTable";

            // Configure export options to preserve column headers and original data types
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,   // first row becomes DataTable column names
                ExportAsString = false,    // keep original .NET types
                CheckMixedValueType = true // verify mixed types and fallback to string if needed
            };

            // Export the ListObject's data range to a DataTable using the options above
            DataTable dataTable = table.DataRange.ExportDataTable(exportOptions);

            // Write the DataTable to a CSV file while preserving data types in the output format
            string csvFilePath = "QueryTableExport.csv";
            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                // Write CSV header
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    writer.Write(dataTable.Columns[col].ColumnName);
                    if (col < dataTable.Columns.Count - 1) writer.Write(",");
                }
                writer.WriteLine();

                // Write each DataRow
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        object value = row[col];

                        // Preserve formatting for dates and numeric types
                        if (value is DateTime dtValue)
                        {
                            // ISO 8601 format for dates
                            writer.Write(dtValue.ToString("o", CultureInfo.InvariantCulture));
                        }
                        else if (value is IFormattable fmtValue)
                        {
                            // Use invariant culture for numbers to avoid locale issues
                            writer.Write(fmtValue.ToString(null, CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            writer.Write(value?.ToString() ?? string.Empty);
                        }

                        if (col < dataTable.Columns.Count - 1) writer.Write(",");
                    }
                    writer.WriteLine();
                }
            }

            // Save the workbook (optional, for verification)
            workbook.Save("WorkbookWithQueryTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}