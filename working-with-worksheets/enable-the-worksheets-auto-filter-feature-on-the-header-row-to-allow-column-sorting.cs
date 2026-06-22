using System;
using Aspose.Cells;

namespace AutoFilterExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data with a header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(15);

            // Enable AutoFilter on the header row (covers the whole data range)
            worksheet.AutoFilter.Range = "A1:B4";

            // Indicate that the range has headers so sorting works correctly
            workbook.DataSorter.HasHeaders = true;

            // Save the workbook
            workbook.Save("AutoFilterEnabled.xlsx");
        }
    }
}