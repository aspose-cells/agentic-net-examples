using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsNumericFormattingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook sourceWorkbook = new Workbook("input.xlsx");
            Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

            // Export the worksheet data to a DataTable (including column headers)
            int maxRow = sourceWorksheet.Cells.MaxDataRow;
            int maxCol = sourceWorksheet.Cells.MaxDataColumn;
            DataTable dataTable = sourceWorksheet.Cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, true);

            // Create a new workbook where the data will be imported with numeric formatting
            Workbook targetWorkbook = new Workbook();
            Worksheet targetWorksheet = targetWorkbook.Worksheets[0];

            // Define import options:
            // - Convert string values that look like numbers to numeric types
            // - Apply the numeric format "0.00" to all columns (you can customize per column)
            ImportTableOptions importOptions = new ImportTableOptions
            {
                ConvertNumericData = true,
                IsFieldNameShown = true,
                // Apply "0.00" format to each column; for simplicity we set the same format for all columns
                NumberFormats = new string[dataTable.Columns.Count]
            };

            // Fill the NumberFormats array with the desired pattern
            for (int i = 0; i < importOptions.NumberFormats.Length; i++)
            {
                importOptions.NumberFormats[i] = "0.00"; // equivalent to {{field_name:0.00}} pattern
            }

            // Import the DataTable into the target worksheet using the defined options
            targetWorksheet.Cells.ImportData(dataTable, 0, 0, importOptions);

            // Save the resulting workbook
            targetWorkbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}