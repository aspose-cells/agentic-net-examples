using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all external links (including hyperlinks) so they are not exported
        workbook.Worksheets.ExternalLinks.Clear(true);

        // Optionally clear any internal hyperlink collections to be extra safe
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Hyperlinks.Clear();
        }

        // Save the workbook as CSV; hyperlinks will not appear in the output
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}