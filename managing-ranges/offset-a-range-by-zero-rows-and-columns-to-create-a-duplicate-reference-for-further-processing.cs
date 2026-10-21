// Title: How to duplicate an Excel range by using CreateRange with zero offset in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# duplicate a range without moving it | Create a copy of an existing range using CreateRange with same size | Offset a range by 0 rows and 0 columns to get a duplicate reference | How to get a second reference to the same cell block in Aspose.Cells | Clone Excel range object in .NET with Aspose.Cells API
// Tags: Aspose.Cells CreateRange zero offset | duplicate range reference Aspose.Cells | range cloning using Aspose.Cells API | Excel range copy without data duplication | C# Aspose.Cells range manipulation

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells A1‑B2, defines a source range, then obtains a second Range object by calling Worksheet.Cells.CreateRange with the source range's first row, first column, row count and column count—effectively an offset of zero rows and columns. The duplicate range is used to change a cell value and the file is saved as OffsetDuplicate.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["B2"].PutValue(4);

            // Create the original range A1:B2
            Aspose.Cells.Range originalRange = sheet.Cells.CreateRange("A1:B2");

            // Duplicate the range by creating a new range with the same dimensions
            Aspose.Cells.Range duplicateRange = sheet.Cells.CreateRange(
                originalRange.FirstRow,
                originalRange.FirstColumn,
                originalRange.RowCount,
                originalRange.ColumnCount);

            // Modify cell A1 through the duplicate range
            duplicateRange[0, 0].PutValue(99);

            // Save the workbook
            workbook.Save("OffsetDuplicate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
