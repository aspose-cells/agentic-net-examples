using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells X1:Y3 (zero‑based indices: row 0, column 23, 3 rows, 2 columns)
        cells.Merge(0, 23, 3, 2);

        // Create a style with the custom number format '#,##0.00 €'
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0.00 €";

        // Apply the style to the merged cell (upper‑left cell X1)
        cells[0, 23].SetStyle(customStyle);

        // Optional: put a numeric value to demonstrate the format
        cells[0, 23].PutValue(1234.56);

        // Save the workbook as ODS using OdsSaveOptions
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        workbook.Save("output.ods", saveOptions);
    }
}