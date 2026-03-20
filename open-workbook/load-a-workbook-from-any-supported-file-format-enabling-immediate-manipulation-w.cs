using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source file – can be any format supported by Aspose.Cells (e.g., .xlsx, .xls, .csv, .pdf, etc.)
        string sourceFile = "sample.xlsx";

        // Load the workbook. The constructor automatically detects the file format.
        Workbook workbook = new Workbook(sourceFile);

        // Access the first worksheet for immediate manipulation.
        Worksheet worksheet = workbook.Worksheets[0];

        // Example manipulation: write a value into cell A1.
        worksheet.Cells["A1"].PutValue("Workbook loaded and modified");

        // Display some basic information to verify the load.
        Console.WriteLine($"Worksheet name: {worksheet.Name}");
        Console.WriteLine($"Cell A1 value: {worksheet.Cells["A1"].StringValue}");

        // Save the modified workbook to a new file (using the Save method).
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}