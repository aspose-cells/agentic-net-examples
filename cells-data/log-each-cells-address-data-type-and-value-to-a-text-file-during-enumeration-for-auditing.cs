// Title: C# – Log cell address, type, and value to a text file while enumerating Aspose.Cells worksheet
// Description: Creates a workbook, adds sample data, then uses the Cells enumerator to walk every cell. For each cell it captures the address (e.g., A1), the data type, and the value, writing them as tab‑separated rows to a text file before saving the workbook.
// Keywords: Aspose.Cells enumerate cells | log cell address C# | export cell type to file | worksheet audit trail .NET | write cell values text file
// Common Searches: Aspose.Cells log each cell address and value | C# enumerate worksheet cells and save to txt | how to audit cell data type with Aspose.Cells | export Aspose.Cells cell metadata to a file
// Developer Intent: Generate a plain‑text audit log that records every cell’s address, data type, and value during worksheet enumeration.
// Use Cases: Compliance reporting: capture a snapshot of all cell contents before distribution. | Debugging: create a readable dump of cell information to identify unexpected data. | Data pipelines: track transformations by logging cell‑level details during processing.
// AI Prompts: Write C# code using Aspose.Cells to enumerate all cells in a worksheet and output address, type, and value to a CSV file. | Show how to extend the logger to include row/column indices and safely handle null values. | Explain how to embed this cell‑audit routine into a larger import workflow with error handling and resource cleanup.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds sample data, then uses the Cells enumerator to walk every cell. For each cell it captures the address (e.g., A1), the data type, and the value, writing them as tab‑separated rows to a text file before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (creation rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (can be replaced with loading logic)
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue(123);
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Path for the audit log file
        string logFilePath = "CellAuditLog.txt";

        // Open a StreamWriter to write the audit information
        using (StreamWriter writer = new StreamWriter(logFilePath, false))
        {
            // Get the cells enumerator (enumeration rule)
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();

            // Iterate through each cell in the worksheet
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Retrieve cell address, data type, and value
                string address = cell.Name;                     // e.g., "A1"
                string dataType = cell.Type.ToString();         // e.g., "IsString"
                string value = cell.Value != null ? cell.Value.ToString() : "null";

                // Write the information to the log file
                writer.WriteLine($"{address}\t{dataType}\t{value}");
            }
        }

        // Save the workbook (save rule)
        workbook.Save("AuditedWorkbook.xlsx");
    }
}
