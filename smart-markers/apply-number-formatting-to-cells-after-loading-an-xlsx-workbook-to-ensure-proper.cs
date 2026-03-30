using System;
using Aspose.Cells;

class ApplyNumberFormatting
{
    static void Main()
    {
        // STEP 1: Create a new workbook (instead of loading input.xlsx)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // STEP 2: Add sample data
        Cells cells = worksheet.Cells;

        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Salary");

        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(1234.567);

        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(9876.5);

        cells["A4"].PutValue("Charlie");
        cells["B4"].PutValue(5000);

        // STEP 3: Create number style
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Custom = "0.00";

        // STEP 4: Apply formatting to numeric cells
        if (cells.MaxDataRow < 0 || cells.MaxDataColumn < 0)
        {
            Console.WriteLine("Worksheet is empty. No formatting applied.");
            return;
        }

        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];

                if (cell.Value is double || cell.Value is int || cell.Value is float || cell.Value is decimal)
                {
                    cell.SetStyle(numberStyle);
                }
            }
        }

        // STEP 5: Save result
        workbook.Save("output.xlsx");

        Console.WriteLine("Sample workbook created and formatting applied successfully.");
    }
}