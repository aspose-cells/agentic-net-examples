using System.Drawing;
using Aspose.Cells;

class PreserveBordersDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Source cell with style and borders ----------
        Cell sourceCell = sheet.Cells["A1"];
        sourceCell.PutValue("Source");

        // Create a style for the source cell
        Style sourceStyle = workbook.CreateStyle();
        sourceStyle.Font.Name = "Arial";
        sourceStyle.Font.Size = 12;
        sourceStyle.Font.IsBold = true;
        sourceStyle.ForegroundColor = Color.LightYellow;
        sourceStyle.Pattern = BackgroundType.Solid;

        // Define borders on the source style
        sourceStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
        sourceStyle.Borders[BorderType.TopBorder].Color = Color.Red;
        sourceStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        sourceStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
        sourceStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        sourceStyle.Borders[BorderType.LeftBorder].Color = Color.Green;
        sourceStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
        sourceStyle.Borders[BorderType.RightBorder].Color = Color.Orange;

        // Apply the style to the source cell
        sourceCell.SetStyle(sourceStyle);

        // ---------- Destination cell ----------
        Cell destCell = sheet.Cells["B2"];
        destCell.PutValue("Destination");

        // Copy the style from source to destination, preserving borders
        Style destStyle = workbook.CreateStyle();
        destStyle.Copy(sourceStyle); // copies all properties, including borders
        destCell.SetStyle(destStyle);

        // ---------- Alternative using Range.CopyStyle (optional) ----------
        // Range srcRange = sheet.Cells.CreateRange("A1");
        // Range destRange = sheet.Cells.CreateRange("C3");
        // destRange.CopyStyle(srcRange); // copies style with borders

        // Save the workbook
        workbook.Save("PreserveBorders.xlsx");
    }
}