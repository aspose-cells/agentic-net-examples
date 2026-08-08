// Title: Audit Worksheet Cells to CSV – Log Address, Value, and Data Type with Aspose.Cells for .NET (C#)
// Description: A C# example that creates a workbook, fills cells with text, numbers, dates and booleans, then enumerates every cell in the first worksheet using Aspose.Cells. For each cell it writes a CSV line containing the cell address (Name), the escaped value, and the Cell.Type enumeration, producing a ready‑to‑use audit file.
// Keywords: Aspose.Cells | C# .NET | cell enumeration | CSV audit log | cell address | cell value | cell data type | Workbook | Worksheet | Cell.Type | export to CSV | data validation | audit trail | forensic spreadsheet
// Common Searches: Aspose.Cells write cell address to CSV | C# enumerate worksheet cells Aspose | Create audit log of Excel cells using Aspose.Cells | Export cell type information to CSV | How to log Excel cell values with Aspose.Cells .NET
// Developer Intent: Generate a CSV file that records each cell’s address, displayed value, and data type while iterating through a worksheet.
// Use Cases: Maintain a change‑tracking log for spreadsheet debugging or validation. | Produce an inventory of cell data types before migrating data to another platform. | Create a compliance‑oriented audit trail for financial or regulatory reporting.
// AI Prompts: Write C# code with Aspose.Cells that enumerates all cells in a worksheet and writes their address, escaped value, and Cell.Type to a CSV file. | Provide a reusable method that accepts a Workbook and a CSV path, handling commas and double quotes correctly in the output. | Explain how to extend the CSV audit to include each cell’s formula and style information.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAuditExample
{
    // A C# example that creates a workbook, fills cells with text, numbers, dates and booleans, then enumerates every cell in the first worksheet using Aspose.Cells. For each cell it writes a CSV line containing the cell address (Name), the escaped value, and the Cell.Type enumeration, producing a ready‑to‑use audit file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Quantity");
            cells["C1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(0.75);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(5);
            cells["C3"].PutValue(0.50);
            cells["A4"].PutValue(DateTime.Now);
            cells["B4"].PutValue(true);

            // Path for the audit CSV file
            string csvPath = "CellAuditLog.csv";

            // Write header line
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                writer.WriteLine("CellAddress,CellValue,CellDataType");

                // Enumerate all cells in the worksheet
                IEnumerator enumerator = cells.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Prepare CSV-safe value (escape double quotes)
                    string valueText = cell.Value?.ToString().Replace("\"", "\"\"") ?? string.Empty;

                    // Write a line with address, value, and data type
                    writer.WriteLine($"{cell.Name},\"{valueText}\",{cell.Type}");
                }
            }

            Console.WriteLine($"Audit log written to {Path.GetFullPath(csvPath)}");
        }
    }
}
