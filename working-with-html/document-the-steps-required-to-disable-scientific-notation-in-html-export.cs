using Aspose.Cells;
using Aspose.Cells.Saving;

class DisableScientificNotationHtmlExport
{
    static void Main()
    {
        // 1. Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // 2. Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // 3. Insert a large numeric value that would normally be displayed in scientific notation
        sheet.Cells["A1"].PutValue(123456789012345);

        // 4. Create a style with a custom number format that forces plain numeric display (no scientific notation)
        Style plainNumberStyle = workbook.CreateStyle();
        plainNumberStyle.Custom = "0"; // Use a simple integer format
        sheet.Cells["A1"].SetStyle(plainNumberStyle);

        // 5. Prepare HTML save options (default options are sufficient for this scenario)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // 6. Save the workbook as an HTML file; the number will appear without scientific notation
        workbook.Save("output.html", htmlOptions);
    }
}