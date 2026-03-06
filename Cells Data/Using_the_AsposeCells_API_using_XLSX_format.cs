using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses the provided constructor)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to cells
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Save the workbook to an XLSX file (uses the provided Save method with SaveFormat)
        workbook.Save("sample.xlsx", SaveFormat.Xlsx);
    }
}