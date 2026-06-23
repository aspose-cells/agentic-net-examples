using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Add a named range that refers to the entire column Z (column Z = 26th column)
        int nameIndex = workbook.Worksheets.Names.Add("ColumnZData");
        // Set the RefersTo formula to the whole column Z on Sheet1
        workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$Z:$Z";

        // (Optional) Retrieve the range to verify
        // Name namedRange = workbook.Worksheets.Names[nameIndex];
        // Range range = namedRange.GetRange();
        // Console.WriteLine($"Named range address: {range.Address}"); // Expected: Z:Z

        // Save the workbook (save rule)
        workbook.Save("ColumnZNamedRange.xlsx");
    }
}