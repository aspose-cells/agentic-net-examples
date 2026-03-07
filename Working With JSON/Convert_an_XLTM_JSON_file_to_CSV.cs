using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Paths
        string sourcePath = "workbook.json";
        string destinationPath = "workbook.csv";

        // Ensure the source JSON file exists; if not, create a sample workbook and save as JSON.
        if (!File.Exists(sourcePath))
        {
            // Create a sample workbook with some data.
            var wb = new Workbook();
            var ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Score");
            ws.Cells["A2"].PutValue("Alice");
            ws.Cells["B2"].PutValue(85);
            ws.Cells["A3"].PutValue("Bob");
            ws.Cells["B3"].PutValue(92);

            // Save the workbook as JSON.
            wb.Save(sourcePath, SaveFormat.Json);
        }

        // Convert the JSON workbook to CSV.
        ConversionUtility.Convert(sourcePath, destinationPath);

        Console.WriteLine($"Conversion completed: {sourcePath} → {destinationPath}");
    }
}