using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to demonstrate the conclusion sheet
            sheet.Cells["A1"].PutValue("Metric");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Total Sales");
            sheet.Cells["B2"].PutValue(125000);
            sheet.Cells["A3"].PutValue("Average Price");
            sheet.Cells["B3"].PutValue(15.75);
            sheet.Cells["A4"].PutValue("Units Sold");
            sheet.Cells["B4"].PutValue(8000);

            // Save the workbook to an XLSX file using the Save method with SaveFormat.Xlsx
            workbook.Save("Conclusion.xlsx", SaveFormat.Xlsx);
        }
    }
}