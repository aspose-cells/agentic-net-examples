using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAudit
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (replace with actual data as needed)
            cells["A1"].PutValue("Sample Text");
            cells["B1"].PutValue(12345);
            cells["C1"].PutValue(DateTime.Now);
            cells["D1"].PutValue(true);
            cells["E1"].PutValue(null); // Blank cell

            // Prepare the audit log file
            string logPath = "CellAuditLog.txt";
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Get the enumerator for all existing cells in the worksheet
                IEnumerator enumerator = cells.GetEnumerator();

                // Iterate through each cell
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Build the log entry: address, type, and value
                    string address = cell.Name;                     // e.g., "A1"
                    string type = cell.Type.ToString();             // CellValueType enum name
                    string value = cell.Value == null ? "null" : cell.Value.ToString();

                    // Write the entry to the log file
                    writer.WriteLine($"{address}\t{type}\t{value}");
                }
            }

            // Save the workbook (demonstrates usage of the provided save rule)
            workbook.Save("AuditWorkbook.xlsx");
        }
    }
}