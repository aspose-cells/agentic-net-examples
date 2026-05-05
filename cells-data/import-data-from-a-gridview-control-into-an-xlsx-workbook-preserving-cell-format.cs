using System;
using System.Data;
using Aspose.Cells;

namespace ExportUtility
{
    public class GridViewToExcelExporter
    {
        /// <summary>
        /// Imports the content of a DataTable into an XLSX workbook while preserving column headers.
        /// </summary>
        /// <param name="dataTable">The DataTable containing the data to export.</param>
        /// <param name="outputPath">Full file path where the workbook will be saved (e.g., "C:\\Data\\ExportedGrid.xlsx").</param>
        public void ExportDataTableToExcel(DataTable dataTable, string outputPath)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Write column headers
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
            }

            // Write data rows
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    cells[row + 1, col].PutValue(dataTable.Rows[row][col]);
                }
            }

            // Save the workbook to the specified path
            workbook.Save(outputPath);
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample DataTable
            DataTable dt = new DataTable("Sample");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Rows.Add(1, "Alice");
            dt.Rows.Add(2, "Bob");

            string outputPath = @"C:\Data\ExportedGrid.xlsx";

            GridViewToExcelExporter exporter = new GridViewToExcelExporter();
            exporter.ExportDataTableToExcel(dt, outputPath);

            Console.WriteLine("Export completed.");
        }
    }
}