using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX file with default load options
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range using the A1 style address syntax (e.g., B2:D4)
            AsposeRange dataRange = cells.CreateRange("B2", "D4");

            // Set a value for the entire range
            dataRange.Value = "Sample";

            // Optionally, you could also set a formula that references the range
            // cells["E2"].Formula = "=SUM(B2:D4)";

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}