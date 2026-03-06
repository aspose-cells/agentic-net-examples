using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Directory containing the Excel file
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        Directory.CreateDirectory(dataDir);

        // Full path to the Excel file
        string filePath = Path.Combine(dataDir, "example.xlsx");

        // If the file does not exist, create a simple workbook and save it
        if (!File.Exists(filePath))
        {
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Name = "SampleSheet";
            ws.Cells["A1"].PutValue("Hello");
            ws.Cells["B2"].PutValue(123);
            wb.Save(filePath);
        }

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Output worksheet information to the console
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Number of Cells: " + worksheet.Cells.Count);
    }
}