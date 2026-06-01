using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class UnionRangeClearContentsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate three separate blocks with sample data
            // Block 1: A1:B2
            cells["A1"].PutValue("A1");
            cells["A2"].PutValue("A2");
            cells["B1"].PutValue("B1");
            cells["B2"].PutValue("B2");

            // Block 2: C3:D4
            cells["C3"].PutValue("C3");
            cells["C4"].PutValue("C4");
            cells["D3"].PutValue("D3");
            cells["D4"].PutValue("D4");

            // Block 3: E5:F6
            cells["E5"].PutValue("E5");
            cells["E6"].PutValue("E6");
            cells["F5"].PutValue("F5");
            cells["F6"].PutValue("F6");

            // Create a UnionRange that combines the three blocks
            // Using WorksheetCollection.CreateUnionRange with a comma‑separated address list
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,C3:D4,E5:F6", 0);

            // Clear contents of each individual range within the union
            foreach (AsposeRange r in unionRange.Ranges)
            {
                r.ClearContents();
            }

            // Save the workbook
            workbook.Save("UnionRangeClearContentsDemo.xlsx");
        }
        catch (Exception ex)
        {
            // Log or display the error
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}