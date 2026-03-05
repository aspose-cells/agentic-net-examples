using System;
using Aspose.Cells;

namespace AsposeCellsImportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data to import: headers followed by rows
            object[] data = new object[]
            {
                "Product", "Price", "Quantity",   // Header row
                "Apple",   1.20,   10,
                "Banana",  0.80,   20,
                "Cherry",  2.50,   15
            };

            // Import the object array horizontally starting at cell A1 (row 0, column 0)
            worksheet.Cells.ImportObjectArray(data, 0, 0, false);

            // Save the workbook in XLSX format
            workbook.Save("ImportedData.xlsx", SaveFormat.Xlsx);
        }
    }
}