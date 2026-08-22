// Title: Create an audit log of each cell’s address, data type, and value while enumerating a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates over all non‑empty cells in an Aspose.Cells worksheet and writes the cell address, type, and value as a tab‑separated line to a text file. | Modify the enumeration to include empty cells and output the audit information in CSV format instead of TSV. | Add comprehensive try‑catch handling around the StreamWriter and ensure the workbook is saved only after the audit file is successfully closed.
// Common Searches: how to export cell address, type, and value from Aspose.Cells to a text file in C# | enumerate worksheet cells and create a TSV audit log using Aspose.Cells .NET | C# Aspose.Cells write cell details (address, data type, value) to a log file | log Excel cell metadata during enumeration with Aspose.Cells and StreamWriter | save audit log of populated cells while generating an Excel workbook in C#
// Tags: Aspose.Cells cell enumeration TSV logging | C# write Excel cell address and type to file | StreamWriter based audit of worksheet cells | Aspose.Cells .NET audit log of populated cells | export cell metadata Aspose.Cells C#

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// Iterates through all populated cells in the first worksheet, writes each cell's address, data type, and value as a tab‑separated line to CellAuditLog.txt, and then saves the workbook as AuditWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data (replace with loading logic if needed)
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue(123);
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Open a text file for writing the audit log
        using (StreamWriter writer = new StreamWriter("CellAuditLog.txt"))
        {
            // Get the cells enumerator for the worksheet
            IEnumerator enumerator = worksheet.Cells.GetEnumerator();

            // Iterate through all cells that contain data
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Prepare log information: address, data type, and value
                string address = cell.Name;                                 // e.g., "A1"
                string dataType = cell.Type.ToString();                     // e.g., "IsString"
                string value = cell.Value != null ? cell.Value.ToString() : "null";

                // Write a tab‑separated line to the log file
                writer.WriteLine($"{address}\t{dataType}\t{value}");
            }
        }

        // Save the workbook (uses the provided save rule)
        workbook.Save("AuditWorkbook.xlsx");
    }
}
