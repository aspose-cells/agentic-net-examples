using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string filePath = "input.xlsx";

        // Load the workbook from the file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(filePath);

        // Get the collection of ContentTypeProperty objects
        ContentTypePropertyCollection ctProps = workbook.ContentTypeProperties;

        // Inspect the collection
        if (ctProps.Count == 0)
        {
            Console.WriteLine("No ContentTypeProperties found in the workbook.");
        }
        else
        {
            Console.WriteLine($"Found {ctProps.Count} ContentTypeProperty(s):");
            for (int i = 0; i < ctProps.Count; i++)
            {
                ContentTypeProperty prop = ctProps[i];
                Console.WriteLine($"Name       : {prop.Name}");
                Console.WriteLine($"Value      : {prop.Value}");
                Console.WriteLine($"Type       : {prop.Type}");
                Console.WriteLine($"IsNillable : {prop.IsNillable}");
                Console.WriteLine(new string('-', 30));
            }
        }

        // Clean up
        workbook.Dispose();
    }
}