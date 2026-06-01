using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLS or XLSX)
        string sourcePath = "input.xlsx";

        // Load the workbook using the string constructor
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Display basic information about the worksheet
        Console.WriteLine($"Worksheet Name: {worksheet.Name}");
        Console.WriteLine($"Cell A1 Value: {worksheet.Cells["A1"].StringValue}");

        // Optionally, save a copy of the loaded workbook
        string destinationPath = "copy.xlsx";
        workbook.Save(destinationPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to: {destinationPath}");
    }
}