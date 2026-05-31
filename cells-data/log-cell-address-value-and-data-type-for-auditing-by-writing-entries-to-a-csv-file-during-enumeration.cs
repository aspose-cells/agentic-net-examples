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
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Populate some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(0.69);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.49);
            sheet.Cells["A4"].PutValue(DateTime.Now);
            sheet.Cells["B4"].PutValue(true);

            // Save the workbook (save rule) – optional, just to demonstrate lifecycle usage
            workbook.Save("SampleData.xlsx");

            // Prepare CSV file for audit logging
            string csvPath = "CellAuditLog.csv";
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("CellAddress,Value,DataType");

                // Get enumerator for all cells in the worksheet (enumeration)
                IEnumerator cellEnumerator = sheet.Cells.GetEnumerator();

                while (cellEnumerator.MoveNext())
                {
                    Cell cell = (Cell)cellEnumerator.Current;

                    // Retrieve address, value and type
                    string address = cell.Name; // e.g., "A1"
                    string value = cell.Value?.ToString() ?? string.Empty;
                    string dataType = cell.Type.ToString(); // e.g., "IsString", "IsNumeric"

                    // Escape commas in value if needed
                    if (value.Contains(","))
                    {
                        value = $"\"{value}\"";
                    }

                    // Write a CSV line
                    writer.WriteLine($"{address},{value},{dataType}");
                }
            }

            Console.WriteLine($"Audit log written to '{csvPath}'.");
        }
    }
}