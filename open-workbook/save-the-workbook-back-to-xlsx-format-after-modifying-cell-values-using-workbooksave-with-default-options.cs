using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – modify cells and save as XLSX with default options
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify some cell values
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Save the workbook to XLSX format using default save options
        workbook.Save("ModifiedWorkbook.xlsx");
    }
}