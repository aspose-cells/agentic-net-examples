using System;
using Aspose.Cells;

class SetFirstPageNumberDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Disable automatic first page numbering
        worksheet.PageSetup.IsAutoFirstPageNumber = false;

        // Set the initial page number (e.g., start from page 5)
        worksheet.PageSetup.FirstPageNumber = 5;

        // Save the workbook with the updated page setup
        workbook.Save("FirstPageNumberDemo.xlsx", SaveFormat.Xlsx);
    }
}