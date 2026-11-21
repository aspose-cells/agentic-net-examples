// Title: How to assign an absolute address to a named range using Aspose.Cells for .NET to keep it fixed
// Common Searches: Aspose.Cells C# set named range RefersTo with absolute address to avoid shifting | Create a fixed named range in an Excel workbook using the Aspose.Cells .NET API | Prevent a named range from moving when rows are inserted with Aspose.Cells | Set the RefersTo property to $A$1:$B$10 in Aspose.Cells C# example | Absolute reference for a named range in an Aspose.Cells workbook
// Tags: Aspose.Cells named range RefersTo absolute address | C# lock named range reference Aspose.Cells | Excel workbook fixed named range Aspose.Cells | Aspose.Cells prevent range shift on row insert | named range absolute reference .NET Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, builds an absolute address string such as "DataSheet!$A$1:$B$10", adds a named range called "MyAbsoluteRange", assigns the absolute address to its RefersTo property to lock the reference, and saves the file as NamedRangeAbsoluteAddress.xlsx.
class SetNamedRangeAbsoluteAddress
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Define the range dimensions (e.g., A1:B10)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 10;        // number of rows in the range
            int totalColumns = 2;      // number of columns in the range

            // Build an absolute address string like "DataSheet!$A$1:$B$10"
            string startCell = $"${CellsHelper.ColumnIndexToName(firstColumn)}${firstRow + 1}";
            string endCell = $"${CellsHelper.ColumnIndexToName(firstColumn + totalColumns - 1)}${firstRow + totalRows}";
            string absoluteAddress = $"{sheet.Name}!{startCell}:{endCell}";

            // Add a named range to the workbook
            int nameIndex = workbook.Worksheets.Names.Add("MyAbsoluteRange");
            Name namedRange = workbook.Worksheets.Names[nameIndex];

            // Set the RefersTo property using the absolute address
            namedRange.RefersTo = absoluteAddress;

            // Optional: verify the RefersTo value
            Console.WriteLine($"Named range '{namedRange.Text}' refers to: {namedRange.RefersTo}");

            // Save the workbook to a file
            workbook.Save("NamedRangeAbsoluteAddress.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
