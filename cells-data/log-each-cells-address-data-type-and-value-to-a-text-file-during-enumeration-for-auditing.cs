using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAuditExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Hello");
            cells["B1"].PutValue(123);
            cells["C1"].PutValue(DateTime.Now);
            cells["A2"].PutValue(true);
            cells["B2"].PutValue(null); // Blank cell

            // Prepare the audit file
            string auditFilePath = "CellAudit.txt";
            using (StreamWriter writer = new StreamWriter(auditFilePath, false))
            {
                // Get the enumerator for all cells that contain data
                IEnumerator enumerator = cells.GetEnumerator();

                // Iterate through each cell
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Retrieve cell address, type, and value
                    string address = cell.Name;                     // e.g., "A1"
                    CellValueType type = cell.Type;                 // enum indicating data type
                    object value = cell.Value;                      // actual cell value (may be null)

                    // Write audit information to the text file
                    writer.WriteLine($"{address}\tType: {type}\tValue: {value ?? "null"}");
                }
            }

            // Save the workbook (optional, just to have a file)
            workbook.Save("AuditedWorkbook.xlsx");
        }
    }
}