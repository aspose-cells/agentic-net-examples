using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the target cell and set its value
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("This is a long text that should be justified and wrapped within the cell.");

        // Retrieve the cell's current style
        Style style = cell.GetStyle();

        // Set horizontal alignment to Justify
        style.HorizontalAlignment = TextAlignmentType.Justify;

        // Enable text wrapping
        style.IsTextWrapped = true;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Auto-fit the row height to display wrapped text properly
        worksheet.AutoFitRow(0);

        // Save the workbook
        workbook.Save("JustifyWrapDemo.xlsx");
    }
}