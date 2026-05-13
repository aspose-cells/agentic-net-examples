using System;
using System.Data;
using System.IO;
using Aspose.Cells;

class ExportVisibleRowsToTxt
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        int totalRows = cells.MaxDataRow + 1;
        int totalColumns = cells.MaxDataColumn + 1;

        // Export all data (including hidden rows) with column names
        DataTable allData = cells.ExportDataTable(0, 0, totalRows, totalColumns, true);

        // Create a table that will contain only visible rows
        DataTable visibleData = allData.Clone();

        for (int rowIndex = 0; rowIndex < allData.Rows.Count; rowIndex++)
        {
            // Check if the corresponding worksheet row is hidden
            if (!worksheet.Cells.Rows[rowIndex].IsHidden)
            {
                visibleData.ImportRow(allData.Rows[rowIndex]);
            }
        }

        string outputPath = "visible_data.txt";

        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            // Write column headers
            for (int col = 0; col < visibleData.Columns.Count; col++)
            {
                writer.Write(visibleData.Columns[col].ColumnName);
                if (col < visibleData.Columns.Count - 1) writer.Write('\t');
            }
            writer.WriteLine();

            // Write visible rows
            foreach (DataRow row in visibleData.Rows)
            {
                for (int col = 0; col < visibleData.Columns.Count; col++)
                {
                    writer.Write(row[col]?.ToString() ?? string.Empty);
                    if (col < visibleData.Columns.Count - 1) writer.Write('\t');
                }
                writer.WriteLine();
            }
        }

        Console.WriteLine($"Visible rows exported successfully to '{outputPath}'.");
    }
}