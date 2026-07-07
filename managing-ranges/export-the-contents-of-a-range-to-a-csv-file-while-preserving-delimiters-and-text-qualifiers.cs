using System;
using System.Text;
using Aspose.Cells;

namespace ExportRangeToCsvDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (including commas and quotes to test qualifiers)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Description");
            cells["C1"].PutValue("Price");

            cells["A2"].PutValue("Widget");
            cells["B2"].PutValue("Small, \"light\" widget"); // contains comma and quotes
            cells["C2"].PutValue(19.99);

            cells["A3"].PutValue("Gadget");
            cells["B3"].PutValue("Large gadget"); // simple text
            cells["C3"].PutValue(29.99);

            // Define the range to export (A1:C3)
            CellArea exportArea = new CellArea
            {
                StartRow = 0,   // Row 0 = A1
                EndRow = 2,     // Row 2 = A3
                StartColumn = 0, // Column 0 = A
                EndColumn = 2    // Column 2 = C
            };

            // Configure text save options for CSV export
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                 // Use comma as delimiter
                ExportArea = exportArea,         // Export only the defined range
                QuoteType = TxtValueQuoteType.Normal, // Quote only when needed (preserves qualifiers)
                Encoding = Encoding.UTF8,        // Ensure proper encoding
                TrimLeadingBlankRowAndColumn = true,
                KeepSeparatorsForBlankRow = false
            };

            // Save the selected range to a CSV file
            workbook.Save("ExportedRange.csv", saveOptions);

            Console.WriteLine("Range exported successfully to ExportedRange.csv");
        }
    }
}