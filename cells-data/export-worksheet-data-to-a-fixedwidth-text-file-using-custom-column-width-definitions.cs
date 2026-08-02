using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFixedWidthExport
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create a workbook and populate data --------------------
            Workbook workbook = new Workbook();                     // create workbook (rule)
            Worksheet sheet = workbook.Worksheets[0];              // get first worksheet

            // Sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Salary");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John Doe");
            sheet.Cells["C2"].PutValue(5230.75);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Jane Smith");
            sheet.Cells["C3"].PutValue(6145.00);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Bob Johnson");
            sheet.Cells["C4"].PutValue(4320.5);

            // -------------------- Define custom column widths (in character units) --------------------
            // These widths will be used for padding/truncating when writing the fixed‑width file.
            double[] columnWidths = new double[] { 5, 20, 10 }; // ID:5, Name:20, Salary:10

            // Apply the same widths to the worksheet (optional, just for visual reference)
            sheet.Cells.SetColumnWidth(0, columnWidths[0]); // Column A
            sheet.Cells.SetColumnWidth(1, columnWidths[1]); // Column B
            sheet.Cells.SetColumnWidth(2, columnWidths[2]); // Column C

            // -------------------- Export the range to a DataTable --------------------
            // Export all rows (including header) and three columns.
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, sheet.Cells.MaxDataRow + 1, 3, true);

            // -------------------- Write the DataTable to a fixed‑width text file --------------------
            string outputPath = "FixedWidthExport.txt";
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (DataRow row in dt.Rows)
                {
                    // Build a line with each column padded/truncated to its defined width.
                    string line = "";
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        string cellText = row[i]?.ToString() ?? string.Empty;

                        // Truncate if longer than the column width.
                        if (cellText.Length > columnWidths[i])
                            cellText = cellText.Substring(0, (int)columnWidths[i]);

                        // Pad right with spaces to fill the column width.
                        line += cellText.PadRight((int)columnWidths[i]);
                    }
                    writer.WriteLine(line);
                }
            }

            // -------------------- (Optional) Save the workbook for verification --------------------
            workbook.Save("WorkbookForReference.xlsx"); // save using rule

            Console.WriteLine($"Fixed‑width text file created at: {Path.GetFullPath(outputPath)}");
        }
    }
}