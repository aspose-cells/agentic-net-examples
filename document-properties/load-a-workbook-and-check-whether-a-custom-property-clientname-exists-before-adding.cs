using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Check if the custom property "ClientName" already exists
        bool hasClientName = workbook.CustomDocumentProperties.Contains("ClientName");

        if (!hasClientName)
        {
            // Add the custom property because it does not exist
            workbook.CustomDocumentProperties.Add("ClientName", "Acme Corp");
            Console.WriteLine("Custom property 'ClientName' added.");
        }
        else
        {
            Console.WriteLine("Custom property 'ClientName' already exists.");
        }

        // Save the workbook (overwrites the original or saves to a new file)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}