using System;
using System.Collections.Generic;
using Aspose.Cells;

public class CustomDataLoader
{
    private readonly Workbook _workbook;

    // Load the workbook from an XLSX file using the provided Workbook(string) constructor
    public CustomDataLoader(string xlsxPath)
    {
        _workbook = new Workbook(xlsxPath);
    }

    // Retrieve all non‑null cell values from the first worksheet as a list of strings
    public List<string> GetFirstWorksheetValues()
    {
        var values = new List<string>();
        Worksheet sheet = _workbook.Worksheets[0];

        foreach (Cell cell in sheet.Cells)
        {
            if (cell.Value != null)
            {
                values.Add(cell.StringValue);
            }
        }

        return values;
    }

    // Save the workbook after any processing (uses the provided Workbook.Save(string) method)
    public void Save(string outputPath)
    {
        _workbook.Save(outputPath);
    }
}

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Instantiate the custom loader which loads the workbook
        var loader = new CustomDataLoader(sourcePath);

        // Extract data from the first worksheet
        List<string> data = loader.GetFirstWorksheetValues();

        // Simple processing: display the extracted values
        Console.WriteLine("Values from the first worksheet:");
        foreach (string value in data)
        {
            Console.WriteLine(value);
        }

        // Optionally save the workbook (could be the same or a new file)
        loader.Save("output.xlsx");
    }
}