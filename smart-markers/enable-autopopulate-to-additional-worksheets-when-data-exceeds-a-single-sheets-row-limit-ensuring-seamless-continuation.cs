// Title: How to auto‑populate additional worksheets in an XLSX file when a DataTable exceeds the 1,048,576‑row limit using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that iterates through a DataTable and creates a new worksheet each time the row index reaches 1,048,576, copying the column headers to the new sheet. | Modify an existing Aspose.Cells workbook to monitor the current row count and programmatically insert a fresh worksheet with the same header row when the maximum rows per sheet are reached. | Implement a configurable MaxRowsPerSheet constant and export a multi‑sheet XLSX workbook from a DataTable containing millions of rows, ensuring each sheet starts with the header row.
// Common Searches: Aspose.Cells split large DataTable across multiple worksheets when row limit is reached | C# export millions of rows to Excel using Aspose.Cells with automatic sheet creation | How to add header row to each new sheet in Aspose.Cells workbook | Configure maximum rows per sheet in Aspose.Cells .NET example | Auto‑populate next worksheet in XLSX when current sheet is full Aspose.Cells
// Tags: auto‑populate worksheets Aspose.Cells | split data across multiple XLSX sheets .NET | row limit handling in Aspose.Cells | copy header row to new worksheet Aspose.Cells | export massive DataTable to multi‑sheet Excel C#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AutoPopulateWorksheets
{
    // The example creates a DataTable with over 2.5 million rows, then uses Aspose.Cells for .NET to write the data to an XLSX workbook. When the row count reaches the XLSX limit of 1,048,576, a new worksheet is added automatically, the header row is copied, and writing continues, resulting in a multi‑sheet file.
    class Program
    {
        // Maximum rows per worksheet for XLSX format
        const int MaxRowsPerSheet = 1048576;

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample data source (replace with actual data retrieval)
                DataTable data = GetSampleData();

                // Initialize first worksheet
                int sheetIndex = 0;
                Worksheet sheet = workbook.Worksheets[sheetIndex];
                sheet.Name = $"Sheet{sheetIndex + 1}";

                // Write header row
                WriteHeader(sheet, data);

                // Start writing data from row 1 (0‑based index)
                int currentRow = 1;

                foreach (DataRow row in data.Rows)
                {
                    // If current row exceeds the limit, create a new worksheet
                    if (currentRow >= MaxRowsPerSheet)
                    {
                        sheetIndex++;
                        int newSheetIdx = workbook.Worksheets.Add(); // returns index of new sheet
                        sheet = workbook.Worksheets[newSheetIdx];
                        sheet.Name = $"Sheet{sheetIndex + 1}";
                        WriteHeader(sheet, data);
                        currentRow = 1; // reset to first data row after header
                    }

                    // Populate cells for the current row
                    for (int col = 0; col < data.Columns.Count; col++)
                    {
                        sheet.Cells[currentRow, col].PutValue(row[col]);
                    }

                    currentRow++;
                }

                // Save the workbook
                string outputPath = "AutoPopulatedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Writes column headers to the given worksheet
        static void WriteHeader(Worksheet sheet, DataTable data)
        {
            for (int col = 0; col < data.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(data.Columns[col].ColumnName);
                // Apply bold style to header
                Style style = sheet.Cells[0, col].GetStyle();
                style.Font.IsBold = true;
                sheet.Cells[0, col].SetStyle(style);
            }
        }

        // Generates a sample DataTable with more rows than a single sheet can hold
        static DataTable GetSampleData()
        {
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Value", typeof(double));

            // Create enough rows to span multiple sheets (e.g., 2.5 million rows)
            int totalRows = (int)(MaxRowsPerSheet * 2.5);
            for (int i = 1; i <= totalRows; i++)
            {
                DataRow dr = table.NewRow();
                dr["ID"] = i;
                dr["Name"] = $"Item {i}";
                dr["Value"] = Math.Sin(i);
                table.Rows.Add(dr);
            }

            return table;
        }
    }
}
