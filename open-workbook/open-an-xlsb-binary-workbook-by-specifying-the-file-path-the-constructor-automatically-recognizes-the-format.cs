using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Specify the path to the XLSB workbook.
        string filePath = "sample.xlsb";

        // Open the workbook. The constructor automatically detects the XLSB format.
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: display worksheet name and the value of cell A1.
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);
    }
}