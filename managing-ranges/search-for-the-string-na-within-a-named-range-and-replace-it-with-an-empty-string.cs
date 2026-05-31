using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

class ReplaceNaInNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells, including "#N/A" values
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("#N/A");
            sheet.Cells["A3"].PutValue(123);
            sheet.Cells["A4"].PutValue("#N/A");
            sheet.Cells["A5"].PutValue("Data");

            // Define a named range that covers the cells to be processed (A2:A4)
            int nameIdx = workbook.Worksheets.Names.Add("MyRange");
            Name namedRange = workbook.Worksheets.Names[nameIdx];
            namedRange.RefersTo = $"={sheet.Name}!$A$2:$A$4";

            // Retrieve the Range object for the named range
            AsposeRange range = namedRange.GetRange();

            // Iterate through each cell in the range and replace "#N/A" with an empty string
            foreach (Cell cell in range)
            {
                if (cell.StringValue == "#N/A")
                {
                    cell.PutValue(string.Empty);
                }
            }

            // Save the modified workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}