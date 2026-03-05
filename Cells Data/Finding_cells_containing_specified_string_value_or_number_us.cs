using System;
using Aspose.Cells;

class FindCellsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Apple");
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["A2"].PutValue("Banana");
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue("Apple Pie");
        sheet.Cells["B3"].PutValue(30);

        // Find the first cell that contains the string "Apple"
        Cell previous = null; // start from the beginning
        Cell foundString = sheet.Cells.Find("Apple", previous);
        if (foundString != null)
        {
            Console.WriteLine($"String \"Apple\" found at {foundString.Name}");

            // Highlight the found cell with a red bold font
            Style style = workbook.CreateStyle();
            style.Font.Color = System.Drawing.Color.Red;
            style.Font.IsBold = true;
            foundString.SetStyle(style);
        }

        // Find the cell that contains the number 20 (exact match)
        FindOptions options = new FindOptions
        {
            LookAtType = LookAtType.EntireContent, // match the whole cell content
            LookInType = LookInType.Values          // search in cell values
        };
        Cell foundNumber = sheet.Cells.Find(20, null, options);
        if (foundNumber != null)
        {
            Console.WriteLine($"Number 20 found at {foundNumber.Name}");

            // Highlight the found cell with a blue bold font
            Style style = workbook.CreateStyle();
            style.Font.Color = System.Drawing.Color.Blue;
            style.Font.IsBold = true;
            foundNumber.SetStyle(style);
        }

        // Save the workbook in XLSX format
        workbook.Save("FindResult.xlsx");
    }
}