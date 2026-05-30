using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Load the HTML file into a workbook
        Workbook workbook = new Workbook(htmlPath);

        // Convert any numeric strings in the worksheet to actual numeric values
        workbook.Worksheets[0].Cells.ConvertStringToNumericValue();

        // Create a style with a custom scientific notation format
        Style sciStyle = workbook.CreateStyle();
        sciStyle.Custom = "0.00E+00";

        // Apply the scientific notation style to all numeric cells
        Cells cells = workbook.Worksheets[0].Cells;
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    cell.SetStyle(sciStyle);
                }
            }
        }

        // Save the workbook as an Excel file
        string excelPath = "output.xlsx";
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}