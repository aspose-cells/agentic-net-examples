using System;
using Aspose.Cells;

class AutoFitColumnsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data representing a simple table
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Description");
        worksheet.Cells["C1"].PutValue("Price");

        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Short item");
        worksheet.Cells["C2"].PutValue(9.99);

        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("This is a much longer description that should cause column B to expand");
        worksheet.Cells["C3"].PutValue(123.45);

        // Auto‑fit all columns so each column width matches its longest cell content
        worksheet.AutoFitColumns();

        // Save the workbook to a file
        workbook.Save("AutoFitColumnsResult.xlsx");
    }
}