using System;
using Aspose.Cells;

class LeadingApostropheHandler
{
    static void Main()
    {
        // Load the XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Iterate through the used range of cells
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Determine if the cell starts with a leading apostrophe
                if (cell.GetStyle().QuotePrefix)
                {
                    // Output cell address and its displayed value
                    Console.WriteLine($"Cell {cell.Name} has leading apostrophe. Displayed value: {cell.StringValue}");
                }
            }
        }

        // Save the workbook to verify that QuotePrefix information is retained
        workbook.Save("output_preserved.xlsx", SaveFormat.Xlsx);
    }
}