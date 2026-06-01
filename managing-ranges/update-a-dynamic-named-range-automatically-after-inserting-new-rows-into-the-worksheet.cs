using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class UpdateDynamicNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with initial data (5 rows)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1);
            }

            // Define a dynamic named range using OFFSET and COUNTA
            // The range will automatically expand/contract based on non‑empty cells in column A
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

            // Display the initial range address resolved by GetRange
            AsposeRange initialRange = dynamicName.GetRange(true);
            Console.WriteLine("Initial range address: " + initialRange.Address);

            // Insert three new rows at index 2 (third row) and update references
            cells.InsertRows(2, 3, true);

            // Add data to the newly inserted rows so the dynamic range should grow
            cells[2, 0].PutValue(100);
            cells[3, 0].PutValue(101);
            cells[4, 0].PutValue(102);

            // Refresh dynamic array formulas (named ranges based on formulas are refreshed here)
            workbook.RefreshDynamicArrayFormulas(true);

            // Retrieve and display the updated range address after insertion
            AsposeRange updatedRange = dynamicName.GetRange(true);
            Console.WriteLine("Updated range address after inserting rows: " + updatedRange.Address);

            // Save the workbook to verify the result
            string outputPath = "DynamicNamedRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}