// Title: Automatically create additional worksheets when a DataTable exceeds Excel’s 1,048,576‑row limit with Aspose.Cells for .NET
// Common Searches: c# aspose.cells split large DataTable across multiple worksheets when row count exceeds Excel limit | how to auto‑add new sheets in an Excel workbook using Aspose.Cells after reaching max rows | populate Excel file from DataTable with automatic worksheet overflow in .NET | write headers to each new worksheet while exporting massive data with Aspose.Cells | handle Excel 2007+ row limit programmatically using Aspose.Cells in C#
// Tags: Aspose.Cells automatic worksheet creation on row limit | export DataTable to multiple Excel sheets .NET | write column headers to each worksheet Aspose.Cells | manage Excel max rows 1048576 with Aspose.Cells | dynamic sheet addition based on data size C# | split large dataset across worksheets using Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AutoPopulateWorksheets
{
    // The example demonstrates how to use Aspose.Cells for .NET to write a large DataTable to an Excel workbook, automatically adding new worksheets when the 1,048,576‑row limit per sheet is reached. It writes column headers on each sheet, names sheets sequentially (Data_1, Data_2, …), and saves the result as an .xlsx file.
    class Program
    {
        // Maximum rows per Excel worksheet (Excel 2007+)
        const int MaxRowsPerSheet = 1048576;

        static void Main(string[] args)
        {
            try
            {
                // ---------- Create a new workbook ----------
                // (Lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data_1";

                // Sample data source: a DataTable with many rows
                DataTable data = GenerateSampleData(2500000); // 2.5 million rows

                // Write data to worksheets, auto‑creating new sheets when row limit is reached
                PopulateWorksheets(workbook, data);

                // ---------- Save the workbook ----------
                // (Lifecycle rule: save)
                string outputPath = "AutoPopulatedWorkbook.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <param name="workbook">The Aspose.Cells workbook.</param>
        /// <param name="data">Data to write.</param>
        static void PopulateWorksheets(Workbook workbook, DataTable data)
        {
            int sheetIndex = 0;               // Current worksheet index
            int currentRow = 0;               // Row within the current worksheet (0‑based)

            // Write column headers once per sheet
            WriteHeaders(workbook.Worksheets[sheetIndex], data);

            // Start writing data rows
            foreach (DataRow dr in data.Rows)
            {
                // If we have reached the maximum rows, create a new sheet
                if (currentRow >= MaxRowsPerSheet - 1) // -1 because row 0 holds headers
                {
                    sheetIndex++;

                    // Worksheets.Add() returns the index of the new sheet in older API versions.
                    int newSheetIdx = workbook.Worksheets.Add();
                    Worksheet newSheet = workbook.Worksheets[newSheetIdx];
                    newSheet.Name = $"Data_{sheetIndex + 1}";

                    WriteHeaders(newSheet, data);
                    currentRow = 0; // Reset for the new sheet (header will be row 0)
                }

                // Write the current data row
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    // Cells are addressed as (row, column) – both zero‑based
                    workbook.Worksheets[sheetIndex].Cells[currentRow + 1, col].PutValue(dr[col]);
                }

                currentRow++;
            }
        }

        /// <summary>
        /// Writes column headers to the specified worksheet.
        /// </summary>
        /// <param name="sheet">Target worksheet.</param>
        /// <param name="data">Data source containing column names.</param>
        static void WriteHeaders(Worksheet sheet, DataTable data)
        {
            for (int col = 0; col < data.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(data.Columns[col].ColumnName);
                // Optional: make header bold
                Style style = sheet.Cells[0, col].GetStyle();
                style.Font.IsBold = true;
                sheet.Cells[0, col].SetStyle(style);
            }
        }

        /// <summary>
        /// Generates a sample DataTable with the specified number of rows.
        /// </summary>
        /// <param name="rowCount">Number of rows to generate.</param>
        /// <returns>Populated DataTable.</returns>
        static DataTable GenerateSampleData(int rowCount)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Value", typeof(double));

            for (int i = 1; i <= rowCount; i++)
            {
                dt.Rows.Add(i, $"Item_{i}", Math.Sin(i));
            }

            return dt;
        }
    }
}
