using System;
using Aspose.Cells;

class AutoFilterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["A5"].PutValue("B");
        worksheet.Cells["B5"].PutValue(40);

        // Apply an auto‑filter to the range that includes the header and data
        worksheet.AutoFilter.Range = "A1:B5";

        // Optional: apply a filter on the first column (Category) for value "A"
        worksheet.AutoFilter.Filter(0, "A");
        worksheet.AutoFilter.Refresh();

        // Save the workbook with the auto‑filter applied
        workbook.Save("AutoFilterDemo.xlsx");
    }
}