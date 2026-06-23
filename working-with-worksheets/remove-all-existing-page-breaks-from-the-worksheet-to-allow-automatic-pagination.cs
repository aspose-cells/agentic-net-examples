using System;
using Aspose.Cells;

class RemovePageBreaksDemo
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet from which page breaks should be removed
        Worksheet worksheet = workbook.Worksheets[0];

        // Clear all horizontal page breaks
        worksheet.HorizontalPageBreaks.Clear();

        // Clear all vertical page breaks
        worksheet.VerticalPageBreaks.Clear();

        // Save the workbook; it will now paginate automatically
        workbook.Save("output.xlsx");
    }
}