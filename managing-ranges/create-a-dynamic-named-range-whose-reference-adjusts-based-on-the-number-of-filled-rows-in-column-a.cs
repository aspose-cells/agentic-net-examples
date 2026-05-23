using System;
using System.IO;
using Aspose.Cells;

class DynamicNamedRangeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with sample data (filled rows)
            for (int i = 0; i < 7; i++) // example: 7 filled rows
            {
                cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Define a dynamic named range that expands based on filled rows in column A
            // Formula uses OFFSET with COUNTA to count non‑empty cells in column A
            // =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$1,0,0,COUNTA({sheet.Name}!$A:$A),1)";

            // Retrieve the range and display its properties
            Aspose.Cells.Range dynRange = dynamicName.GetRange();
            Console.WriteLine($"Dynamic range address: {dynRange.RefersTo}");
            Console.WriteLine($"Rows in dynamic range: {dynRange.RowCount}");
            Console.WriteLine($"Columns in dynamic range: {dynRange.ColumnCount}");

            // Save the workbook
            string outputPath = "DynamicNamedRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}