using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSX workbook
        string filePath = "input.xlsx";

        // Load the workbook using the constructor that accepts a file path
        Workbook workbook = new Workbook(filePath);

        // Get the built‑in document properties collection
        var builtInProps = workbook.BuiltInDocumentProperties;

        // Define the introductory properties you want to display
        string[] introPropertyNames = new string[]
        {
            "Title",
            "Subject",
            "Author",
            "Keywords",
            "Comments",
            "Template",
            "Last Author",
            "Creation Date",
            "Last Save Time",
            "Company"
        };

        // Iterate through the selected properties and output their values
        foreach (string propName in introPropertyNames)
        {
            try
            {
                var prop = builtInProps[propName];
                Console.WriteLine($"{propName}: {prop.Value}");
            }
            catch
            {
                // Property not present – ignore or handle as needed
            }
        }

        // Clean up resources
        workbook.Dispose();
    }
}