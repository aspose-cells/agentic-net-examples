using System;
using Aspose.Cells;

class OpenExcel97To2003
{
    static void Main()
    {
        // Path to the existing Excel 97‑2003 workbook (.xls)
        string filePath = "sample.xls";

        // Load the workbook using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Display basic information
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Number of Cells: " + worksheet.Cells.Count);

        // Save a copy of the workbook (still in Excel 97‑2003 format)
        workbook.Save("copy_of_sample.xls", SaveFormat.Excel97To2003);
    }
}