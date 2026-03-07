using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the XLSX workbook (template containing smart markers)
        Workbook workbook = new Workbook("template.xlsx");

        // Initialize WorkbookDesigner to process smart markers
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Example data source with a DateTime value for the smart marker
        var data = new { Date = DateTime.Now };
        designer.SetDataSource("Data", data);

        // Process all smart markers in the workbook
        designer.Process();

        // After processing, format every cell that contains a DateTime value
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsDateTime)
                {
                    Style style = cell.GetStyle();
                    style.Custom = "MM/dd/yyyy";   // Desired date format
                    cell.SetStyle(style);
                }
            }
        }

        // Save the workbook with formatted dates
        workbook.Save("output.xlsx");
    }
}