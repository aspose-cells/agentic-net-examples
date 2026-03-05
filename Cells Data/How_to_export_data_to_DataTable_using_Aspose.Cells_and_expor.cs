using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsExportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including header row)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("City");

            cells["A2"].PutValue("John");
            cells["B2"].PutValue(28);
            cells["C2"].PutValue("New York");

            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(34);
            cells["C3"].PutValue("London");

            cells["A4"].PutValue("Bob");
            cells["B4"].PutValue(45);
            cells["C4"].PutValue("Sydney");

            // Export the range (including header) to a DataTable
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, exportColumnName
            DataTable dt = cells.ExportDataTable(0, 0, 4, 3, true);

            // Display exported DataTable content
            Console.WriteLine("Exported DataTable:");
            foreach (DataRow row in dt.Rows)
            {
                Console.WriteLine($"{row["Name"]}, {row["Age"]}, {row["City"]}");
            }

            // Save the workbook to XLSX format
            workbook.Save("ExportedData.xlsx");
        }
    }
}