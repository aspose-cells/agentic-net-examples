using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing Excel file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Convert all string-formatted numeric (and date) values in the first worksheet to true numeric values
        workbook.Worksheets[0].Cells.ConvertStringToNumericValue();

        // Save the modified workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}