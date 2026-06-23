using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a sample numeric value
        sheet.Cells["A1"].PutValue(12345.6789);

        // Define a style with scientific notation (three significant digits)
        // Number = 11 corresponds to the built‑in format "0.00E+00"
        Style sciStyle = workbook.CreateStyle();
        sciStyle.Number = 11;

        // Apply only the number format part of the style
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to cell A1
        sheet.Cells.CreateRange("A1").ApplyStyle(sciStyle, flag);

        // Save the workbook
        workbook.Save("ScientificNotationThreeDigits.xlsx");
    }
}