using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the custom document properties collection
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Check if the property "ClientName" already exists
        if (!customProps.Contains("ClientName"))
        {
            // Property does not exist, add it with a sample value
            customProps.Add("ClientName", "Acme Corp");
            Console.WriteLine("Custom property 'ClientName' added.");
        }
        else
        {
            Console.WriteLine("Custom property 'ClientName' already exists.");
        }

        // Save the workbook with the (potentially) new property
        workbook.Save("output.xlsx");
    }
}