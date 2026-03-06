using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Display each custom property: name, value, and data type
        Console.WriteLine("Custom Document Properties:");
        foreach (DocumentProperty prop in customProps)
        {
            Console.WriteLine($"{prop.Name}: {prop.Value} ({prop.Type})");
        }

        // Example: add a new custom property and save the workbook
        customProps.Add("ReviewedBy", "Alice");
        workbook.Save("output.xlsx");
    }
}