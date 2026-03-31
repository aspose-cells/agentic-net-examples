using System;
using Aspose.Cells;

namespace HideHeadersDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide row and column headers for a cleaner presentation
            worksheet.IsRowColumnHeadersVisible = false;

            // Optional: add some sample data to verify the workbook content
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("Data");

            // Save the workbook
            workbook.Save("HiddenHeaders.xlsx", SaveFormat.Xlsx);
        }
    }
}