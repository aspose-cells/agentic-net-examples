using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX file with explicit load options
        string inputPath = "input.xlsx";
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Desired date format
        string dateFormat = "MM/dd/yyyy";

        // Apply the date format to all cells that contain DateTime values
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsDateTime)
                {
                    Style style = cell.GetStyle();
                    style.Custom = dateFormat;
                    cell.SetStyle(style);
                }
            }
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}