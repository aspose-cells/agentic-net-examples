using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsAggregationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Number of tables to import
            int tableCount = 3;

            // Loop to create sample data tables and import them into separate worksheets
            for (int i = 0; i < tableCount; i++)
            {
                // Create a sample DataTable with numeric column "Quantity"
                DataTable dt = new DataTable($"Table{i + 1}");
                dt.Columns.Add("Item", typeof(string));
                dt.Columns.Add("Quantity", typeof(int));
                dt.Columns.Add("Price", typeof(double));

                // Add sample rows
                dt.Rows.Add("Apple", 10 + i, 0.5);
                dt.Rows.Add("Banana", 20 + i, 0.3);
                dt.Rows.Add("Cherry", 15 + i, 0.8);

                // Ensure the worksheet exists (Workbook starts with one sheet)
                Worksheet ws;
                if (i < workbook.Worksheets.Count)
                {
                    ws = workbook.Worksheets[i];
                }
                else
                {
                    ws = workbook.Worksheets[workbook.Worksheets.Add()];
                }

                ws.Name = $"Data{i + 1}";

                // Import the DataTable starting at cell A1 (row 0, column 0)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true // include column headers
                };
                ws.Cells.ImportData(dt, 0, 0, importOptions);
            }

            // Add a new worksheet for the summary
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Summary";

            // Write headers in the summary sheet
            summarySheet.Cells[0, 0].PutValue("Source Sheet");
            summarySheet.Cells[0, 1].PutValue("Total Quantity");

            // Iterate over each data worksheet to calculate totals
            for (int i = 0; i < tableCount; i++)
            {
                Worksheet dataWs = workbook.Worksheets[i];
                // Determine the column index of "Quantity" (assumes it is the second column)
                int quantityColIndex = 1; // zero‑based index for column B

                // Find the last row that contains data in the Quantity column
                int lastRow = dataWs.Cells.GetLastDataRow(quantityColIndex);
                double totalQuantity = 0;

                // Sum values starting from row 1 (skip header)
                for (int row = 1; row <= lastRow; row++)
                {
                    totalQuantity += dataWs.Cells[row, quantityColIndex].DoubleValue;
                }

                // Write results to the summary sheet
                int summaryRow = i + 1; // start after header
                summarySheet.Cells[summaryRow, 0].PutValue(dataWs.Name);
                summarySheet.Cells[summaryRow, 1].PutValue(totalQuantity);
            }

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("AggregatedSummary.xlsx", SaveFormat.Xlsx);
        }
    }
}