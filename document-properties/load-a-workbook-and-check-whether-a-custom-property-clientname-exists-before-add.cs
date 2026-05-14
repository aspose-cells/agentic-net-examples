using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the collection of custom document properties
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Check if the custom property "ClientName" already exists
            if (!customProps.Contains("ClientName"))
            {
                // Property does not exist, add it with a sample value
                customProps.Add("ClientName", "Acme Corp");
                Console.WriteLine("Custom property 'ClientName' added.");
            }
            else
            {
                // Property exists, optionally read its value
                DocumentProperty existingProp = customProps["ClientName"];
                Console.WriteLine($"Custom property 'ClientName' already exists with value: {existingProp.Value}");
            }

            // Save the workbook with the (potentially) new property
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}