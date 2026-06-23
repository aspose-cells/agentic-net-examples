using Aspose.Cells;

class SetNarrowMargins
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set narrow margins (values in centimeters)
        sheet.PageSetup.LeftMargin = 0.5;   // 0.5 cm left margin
        sheet.PageSetup.RightMargin = 0.5;  // 0.5 cm right margin
        sheet.PageSetup.TopMargin = 0.5;    // 0.5 cm top margin
        sheet.PageSetup.BottomMargin = 0.5; // 0.5 cm bottom margin

        // Save the workbook (lifecycle: save)
        workbook.Save("NarrowMargins.xlsx");
    }
}