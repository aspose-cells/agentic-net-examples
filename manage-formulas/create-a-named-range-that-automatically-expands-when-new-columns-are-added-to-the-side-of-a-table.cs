using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate a simple table with headers (A1:C4)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Score");

            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(85);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(90);

            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Charlie");
            worksheet.Cells["C4"].PutValue(78);

            // Convert the range into a ListObject (Excel Table)
            int tableIndex = worksheet.ListObjects.Add("A1", "C4", true);
            var table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Create a named range that refers to the whole table using a structured reference.
            int nameIndex = workbook.Worksheets.Names.Add("MyTableRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=MyTable";

            // Show the address of the named range before adding a column
            AsposeRange rangeBefore = workbook.Worksheets.Names[nameIndex].GetRange();
            Console.WriteLine("Named range before adding column: " + rangeBefore.Address);

            // Insert a new column to the right of the existing table (after column C)
            // Column index is zero‑based, so 3 corresponds to column D.
            worksheet.Cells.InsertColumn(3);

            // Add header and sample data for the new column
            worksheet.Cells["D1"].PutValue("Level");
            worksheet.Cells["D2"].PutValue("A");
            worksheet.Cells["D3"].PutValue("B");
            worksheet.Cells["D4"].PutValue("A");

            // Retrieve the named range again; it now covers the expanded table.
            AsposeRange rangeAfter = workbook.Worksheets.Names[nameIndex].GetRange();
            Console.WriteLine("Named range after adding column: " + rangeAfter.Address);

            // Save the workbook (ensure the directory exists)
            string outputPath = "NamedRangeAutoExpand.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}