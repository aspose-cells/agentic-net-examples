using System;
using Aspose.Cells;

public class LoadXlsxDemo
{
    public static void Run()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Create LoadOptions specifying the XLSX format explicitly
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Load the workbook using the constructor that accepts a file path and LoadOptions
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Iterate through all worksheets to demonstrate that data is preserved
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}, Total Cells: {sheet.Cells.Count}");

            // Example: display the value of cell A1 if it contains data
            Cell a1 = sheet.Cells["A1"];
            if (a1 != null && a1.Value != null)
            {
                Console.WriteLine($"A1 Value: {a1.StringValue}");
            }
        }

        // Save the workbook to a new file, preserving all worksheets and cell data
        string outputPath = "output_copy.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook successfully saved to: {outputPath}");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        LoadXlsxDemo.Run();
    }
}