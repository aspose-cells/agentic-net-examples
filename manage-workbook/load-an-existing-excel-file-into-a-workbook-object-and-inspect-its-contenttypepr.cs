using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsContentTypeDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string filePath = "Sample.xlsx";

            // Load the workbook using the string constructor (provided rule)
            Workbook workbook = new Workbook(filePath);

            // Access the collection of ContentTypeProperty objects
            ContentTypePropertyCollection ctProps = workbook.ContentTypeProperties;

            // If there are no content type properties, inform the user
            if (ctProps.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any ContentTypeProperties.");
                return;
            }

            // Iterate through the collection and display each property's details
            for (int i = 0; i < ctProps.Count; i++)
            {
                ContentTypeProperty prop = ctProps[i];
                Console.WriteLine($"Property #{i + 1}");
                Console.WriteLine($"  Name       : {prop.Name}");
                Console.WriteLine($"  Value      : {prop.Value}");
                Console.WriteLine($"  Type       : {prop.Type}");
                Console.WriteLine($"  IsNillable : {prop.IsNillable}");
                Console.WriteLine();
            }

            // No need to save the workbook as we only inspected properties
        }
    }
}