using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Instantiate a new Workbook object (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the first worksheet (optional, demonstrates usage)
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: put a value in cell A1
        worksheet.Cells["A1"].PutValue("Workbook created successfully");
        
        // (Saving is not required for this task, but can be done with workbook.Save if needed)
        Console.WriteLine("Workbook instance created.");
    }
}