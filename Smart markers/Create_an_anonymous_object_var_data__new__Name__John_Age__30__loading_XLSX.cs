using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook (uses the Workbook(string) constructor)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create an anonymous object with Name and Age properties
        var data = new { Name = "John", Age = 30 };

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Write the anonymous object's values into cells
        sheet.Cells["A1"].PutValue(data.Name);
        sheet.Cells["B1"].PutValue(data.Age);

        // Save the modified workbook (uses the Save(string) method)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}