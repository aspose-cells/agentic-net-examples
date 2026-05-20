using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the custom document properties collection
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        Console.WriteLine("DateTime custom properties:");

        // Iterate through all custom properties and output those of DateTime type
        foreach (DocumentProperty prop in customProps)
        {
            if (prop.Type == PropertyType.DateTime)
            {
                // Display property name and its DateTime value in ISO 8601 format
                Console.WriteLine($"{prop.Name}: {prop.ToDateTime():O}");
            }
        }

        // Save the workbook (optional, as no changes are made to the file)
        workbook.Save("output.xlsx");
    }
}