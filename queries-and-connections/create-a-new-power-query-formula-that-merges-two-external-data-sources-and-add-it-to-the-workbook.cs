// Title: How to merge two CSV files on an ID column and export the combined data to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# example for loading CSV files into DataTables with Aspose.Cells LoadOptions | Aspose.Cells inner join two DataTables on a key column and write to a new worksheet | Export merged DataTable to .xlsx file using Aspose.Cells programmatically | Handle duplicate column names when merging CSV data with Aspose.Cells | Save workbook after importing merged CSV data with Aspose.Cells in .NET
// Tags: Aspose.Cells load CSV to DataTable C# | Aspose.Cells perform inner join on DataTables | Aspose.Cells write DataTable to worksheet cells | Aspose.Cells export merged data to XLSX | Aspose.Cells resolve duplicate column names during merge

using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsPowerQueryExample
{
    // The sample loads two CSV files into DataTables using Aspose.Cells LoadOptions, validates that both contain an 'ID' column, performs an inner join on that column via LINQ, creates a new worksheet named 'MergedData', writes column headers and merged rows into the worksheet cells, and saves the workbook as 'MergedDataWorkbook.xlsx', with error handling for missing files and import issues.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths to the source CSV files
                string sourcePath1 = @"C:\Data\Source1.csv";
                string sourcePath2 = @"C:\Data\Source2.csv";

                // Verify that the source files exist
                if (!File.Exists(sourcePath1))
                    throw new FileNotFoundException($"Source file not found: {sourcePath1}");
                if (!File.Exists(sourcePath2))
                    throw new FileNotFoundException($"Source file not found: {sourcePath2}");

                // Load CSV files into DataTables
                DataTable dt1 = LoadCsvToDataTable(sourcePath1);
                DataTable dt2 = LoadCsvToDataTable(sourcePath2);

                // Merge the two tables on the "ID" column (inner join)
                DataTable mergedTable = MergeDataTables(dt1, dt2);

                // Create a new workbook and add a worksheet for the merged data
                Workbook workbook = new Workbook();
                int sheetIndex = workbook.Worksheets.Add();
                Worksheet mergedWorksheet = workbook.Worksheets[sheetIndex];
                mergedWorksheet.Name = "MergedData";

                // Manually import the DataTable into the worksheet (header + rows)
                try
                {
                    int currentRow = 0;

                    // Write column headers
                    foreach (DataColumn col in mergedTable.Columns)
                    {
                        mergedWorksheet.Cells[currentRow, col.Ordinal].PutValue(col.ColumnName);
                    }

                    // Write data rows
                    currentRow++;
                    foreach (DataRow dr in mergedTable.Rows)
                    {
                        for (int c = 0; c < mergedTable.Columns.Count; c++)
                        {
                            mergedWorksheet.Cells[currentRow, c].PutValue(dr[c]);
                        }
                        currentRow++;
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to import DataTable to worksheet: {ex.Message}", ex);
                }

                // Save the workbook
                string outputPath = "MergedDataWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Loads a CSV file into a DataTable using Aspose.Cells
        private static DataTable LoadCsvToDataTable(string csvPath)
        {
            try
            {
                // Load the CSV file into a temporary workbook
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                Workbook tempWb = new Workbook(csvPath, loadOptions);
                Worksheet tempWs = tempWb.Worksheets[0];

                // Determine the used range
                int maxRow = tempWs.Cells.MaxDataRow;
                int maxCol = tempWs.Cells.MaxDataColumn;

                // Export the range to a DataTable (first row as column names)
                DataTable dt = tempWs.Cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, true);
                return dt;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to load CSV '{csvPath}': {ex.Message}", ex);
            }
        }

        // Performs an inner join on two DataTables based on the "ID" column
        private static DataTable MergeDataTables(DataTable dt1, DataTable dt2)
        {
            // Ensure both tables contain the "ID" column
            if (!dt1.Columns.Contains("ID") || !dt2.Columns.Contains("ID"))
                throw new ArgumentException("Both source tables must contain an 'ID' column.");

            // Prepare the result table schema
            DataTable result = new DataTable("Merged");
            result.Columns.Add("ID", typeof(string));

            // Add columns from the first table (excluding ID)
            foreach (DataColumn col in dt1.Columns)
            {
                if (col.ColumnName == "ID") continue;
                result.Columns.Add(col.ColumnName, typeof(string));
            }

            // Add columns from the second table (excluding ID) with a suffix to avoid duplicates
            foreach (DataColumn col in dt2.Columns)
            {
                if (col.ColumnName == "ID") continue;
                string newColName = col.ColumnName;
                if (result.Columns.Contains(newColName))
                    newColName = col.ColumnName + "_2";
                result.Columns.Add(newColName, typeof(string));
            }

            // Perform the inner join using LINQ
            var query = from row1 in dt1.AsEnumerable()
                        join row2 in dt2.AsEnumerable()
                        on row1.Field<string>("ID") equals row2.Field<string>("ID")
                        select new { Row1 = row1, Row2 = row2 };

            // Populate the result table
            foreach (var pair in query)
            {
                DataRow newRow = result.NewRow();
                newRow["ID"] = pair.Row1["ID"];

                // Copy values from the first table
                foreach (DataColumn col in dt1.Columns)
                {
                    if (col.ColumnName == "ID") continue;
                    newRow[col.ColumnName] = pair.Row1[col.ColumnName];
                }

                // Copy values from the second table
                foreach (DataColumn col in dt2.Columns)
                {
                    if (col.ColumnName == "ID") continue;
                    string targetColName = col.ColumnName;
                    if (result.Columns.Contains(col.ColumnName) && dt1.Columns.Contains(col.ColumnName))
                        targetColName = col.ColumnName + "_2";
                    newRow[targetColName] = pair.Row2[col.ColumnName];
                }

                result.Rows.Add(newRow);
            }

            return result;
        }
    }
}
