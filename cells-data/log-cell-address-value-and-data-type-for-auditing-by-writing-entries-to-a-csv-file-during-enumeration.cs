// Title: Log each cell’s address, value, and data type to a CSV file while enumerating an Aspose.Cells worksheet in C#
// AI Prompts: Generate C# code that iterates through all cells of an Aspose.Cells worksheet and writes the cell address, stringified value, and the cell’s data type into a CSV file with a header row. | Add logic to the enumeration loop that encloses cell values containing commas in double quotes before writing them to the CSV. | Include code that saves the workbook after completing the CSV audit and prints the path of the generated audit file.
// Common Searches: how to export Aspose.Cells worksheet cell details to CSV in C# | C# enumerate cells in Aspose.Cells and write address and type to a log file | escaping commas when writing Excel cell values to CSV using Aspose.Cells | create an audit CSV of cell addresses, values, and data types with Aspose.Cells .NET
// Tags: Aspose.Cells cell enumeration CSV | cell address and type logging .NET | comma handling for CSV output C# | audit file generation from Excel workbook | worksheet cell data type extraction Aspose

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// // Demonstrates creating a workbook, populating sample cells, enumerating every cell, and writing each cell's address, value, and data type to a CSV audit file with proper comma escaping, then saving the workbook.
class CellAuditLogger
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["C1"].PutValue(DateTime.Now);
        cells["A2"].PutValue(true);
        cells["B2"].PutValue(null);
        cells["C2"].PutValue(45.67);

        // Path for the audit CSV file
        string csvPath = "CellAudit.csv";

        // Write audit information to CSV
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // CSV header
            writer.WriteLine("Address,Value,DataType");

            // Enumerate all cells in the worksheet
            IEnumerator enumerator = cells.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;

                // Prepare cell details
                string address = cell.Name;
                string value = cell.Value?.ToString() ?? string.Empty;
                string dataType = cell.Type.ToString();

                // Escape commas in the value field
                if (value.Contains(","))
                {
                    value = $"\"{value}\"";
                }

                // Write a CSV line
                writer.WriteLine($"{address},{value},{dataType}");
            }
        }

        // Save the workbook (optional, demonstrates normal lifecycle)
        workbook.Save("SampleData.xlsx");

        Console.WriteLine($"Cell audit completed. CSV file saved to: {csvPath}");
    }
}
