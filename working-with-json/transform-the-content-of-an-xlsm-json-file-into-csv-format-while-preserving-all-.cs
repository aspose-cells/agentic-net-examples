using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);

        // Path for the resulting CSV file
        string csvPath = "output.csv";

        // Save the workbook content to CSV format
        workbook.Save(csvPath, SaveFormat.Csv);

        Console.WriteLine($"Conversion completed: workbook -> '{csvPath}'");
    }
}