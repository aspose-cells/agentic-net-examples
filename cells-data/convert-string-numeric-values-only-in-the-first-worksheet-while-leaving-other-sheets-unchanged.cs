using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet's cells collection
        Cells firstSheetCells = workbook.Worksheets[0].Cells;

        // Convert all string values that can be interpreted as numbers (or dates) to numeric values
        firstSheetCells.ConvertStringToNumericValue();

        // Save the workbook; other worksheets remain unchanged
        workbook.Save("output.xlsx");
    }
}