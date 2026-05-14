using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the Excel file whose built‑in properties will be read
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Retrieve the collection of built‑in document properties
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Iterate over each property in the collection and output its name and value
        for (int i = 0; i < builtInProps.Count; i++)
        {
            DocumentProperty prop = builtInProps[i];
            Console.WriteLine($"{prop.Name}: {prop.Value}");
        }
    }
}