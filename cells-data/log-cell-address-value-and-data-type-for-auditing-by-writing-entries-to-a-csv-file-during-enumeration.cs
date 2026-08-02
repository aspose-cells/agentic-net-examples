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
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(0.69);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(0.49);
            cells["A4"].PutValue(DateTime.Now);
            cells["B4"].PutValue(true);

            // Path for the audit CSV file
            string csvPath = "audit.csv";

            // Open a StreamWriter for the CSV file
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // Write CSV header
                writer.WriteLine("Address,Value,DataType");

                // Get the enumerator for all cells in the worksheet
                IEnumerator enumerator = cells.GetEnumerator();

                // Iterate through each cell
                while (enumerator.MoveNext())
                {
                    Cell cell = (Cell)enumerator.Current;

                    // Prepare cell value as string; replace commas to avoid CSV column shift
                    string valueString = cell.Value?.ToString().Replace(",", ";") ?? string.Empty;

                    // Write a line with address, value, and the cell's value type
                    writer.WriteLine($"{cell.Name},{valueString},{cell.Type}");
                }
            }

            Console.WriteLine($"Audit completed. CSV file created at: {Path.GetFullPath(csvPath)}");
        }
    }
}