using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnableFilterArrowsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header row (row 1)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["C1"].PutValue("Price");

            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue("Electronics");
            worksheet.Cells["C2"].PutValue(1200);

            worksheet.Cells["A3"].PutValue("Shirt");
            worksheet.Cells["B3"].PutValue("Clothing");
            worksheet.Cells["C3"].PutValue(45);

            worksheet.Cells["A4"].PutValue("Phone");
            worksheet.Cells["B4"].PutValue("Electronics");
            worksheet.Cells["C4"].PutValue(800);

            // Apply an AutoFilter to the data range.
            // The filter arrows will appear on the header row (A1:C1).
            worksheet.AutoFilter.Range = "A1:C4";

            // Protect the worksheet but allow sorting and filtering.
            worksheet.Protect(ProtectionType.All);
            worksheet.Protection.AllowSorting = true;
            worksheet.Protection.AllowFiltering = true;

            // Save the workbook (the file will contain filter arrows on the header row).
            workbook.Save("WorksheetWithFilterArrows.xlsx");
        }
    }
}