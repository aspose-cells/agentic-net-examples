using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (original formatting is default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Modify some data while keeping existing formatting
            sheet.Cells["A1"].PutValue("Original Formatting Preserved");
            sheet.Cells["B2"].PutValue(12345);

            // Save the workbook to XLSX format; this preserves all original formatting
            workbook.Save("ModifiedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook saved as 'ModifiedWorkbook.xlsx' with original formatting.");
        }
    }
}