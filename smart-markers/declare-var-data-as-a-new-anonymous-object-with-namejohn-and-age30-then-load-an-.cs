using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Declare an anonymous object with Name and Age
        var data = new { Name = "John", Age = 30 };
        Console.WriteLine($"Name: {data.Name}, Age: {data.Age}");

        // Load an existing XLSX file using the Workbook(string) constructor
        string filePath = "sample.xlsx"; // replace with your actual file path
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet to demonstrate that the file is loaded
        Worksheet worksheet = workbook.Worksheets[0];
        Console.WriteLine($"First worksheet name: {worksheet.Name}");
    }
}