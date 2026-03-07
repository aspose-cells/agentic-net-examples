using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FirstPageNumberDemo
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Disable automatic first page numbering
            worksheet.PageSetup.IsAutoFirstPageNumber = false;

            // Set the first page number to start from 3
            worksheet.PageSetup.FirstPageNumber = 3;

            // Optional: display the settings in console
            Console.WriteLine("IsAutoFirstPageNumber: " + worksheet.PageSetup.IsAutoFirstPageNumber);
            Console.WriteLine("FirstPageNumber: " + worksheet.PageSetup.FirstPageNumber);

            // Save the workbook to an XLSX file
            workbook.Save("FirstPageNumberDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}