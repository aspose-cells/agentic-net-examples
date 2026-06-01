using System;
using Aspose.Cells;

namespace AsposeCellsHtmlStringDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign HTML markup to cell A1 using the HtmlString property
            // This will display bold and italic text in the cell
            worksheet.Cells["A1"].HtmlString = "This is <b>bold</b> and <i>italic</i> text";

            // Save the workbook to an XLSX file
            workbook.Save("HtmlStringDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}