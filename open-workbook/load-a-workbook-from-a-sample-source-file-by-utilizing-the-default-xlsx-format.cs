using System;
using Aspose.Cells;

class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the sample source file (XLSX format)
        string filePath = "sample.xlsx";

        // Load the workbook using the string constructor.
        // This utilizes the default XLSX format as required.
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the loaded workbook.
        Worksheet worksheet = workbook.Worksheets[0];

        // Example operation: read and display the value of cell A1.
        Console.WriteLine("Cell A1 value: " + worksheet.Cells["A1"].StringValue);
    }
}