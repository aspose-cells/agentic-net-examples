using System.Drawing;
using Aspose.Cells;

class ApplyHeaderRowStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define a style for the header row
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Pattern = BackgroundType.Solid;          // solid fill
        headerStyle.ForegroundColor = Color.LightGray;      // light gray background

        // Create a StyleFlag to apply all style properties
        StyleFlag flag = new StyleFlag();
        flag.All = true;

        // Apply the style to the first row (header row, index 0)
        cells.ApplyRowStyle(0, headerStyle, flag);

        // (Optional) Add some header text to demonstrate the styled row
        cells[0, 0].PutValue("Column A");
        cells[0, 1].PutValue("Column B");
        cells[0, 2].PutValue("Column C");

        // Save the workbook
        workbook.Save("HeaderRowStyled.xlsx");
    }
}