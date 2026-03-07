using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source workbook (change extension to a supported format if needed)
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "input.xlsx");
        string jsonOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.json");

        Workbook workbook;

        if (File.Exists(sourcePath))
        {
            // Load the existing workbook
            workbook = new Workbook(sourcePath);
        }
        else
        {
            // Create a sample workbook if the source file does not exist
            workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Add header row
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Value");

            // Add some data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Item1");
            sheet.Cells["C2"].PutValue(123.45);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Item2");
            sheet.Cells["C3"].PutValue(678.90);
        }

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportNestedStructure = true,
            HasHeaderRow = true,
            ExportEmptyCells = true,
            ToExcelStruct = true
        };

        // Save the workbook as a JSON file using the configured options
        workbook.Save(jsonOutputPath, jsonOptions);

        Console.WriteLine($"Workbook has been converted to JSON at '{jsonOutputPath}'.");
    }
}