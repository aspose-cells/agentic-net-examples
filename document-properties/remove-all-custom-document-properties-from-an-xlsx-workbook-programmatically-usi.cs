using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class RemoveAllCustomDocumentProperties
    {
        public static void Run()
        {
            // Paths to the source and destination workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of custom document properties
            CustomDocumentPropertyCollection customProperties = workbook.CustomDocumentProperties;

            // Collect the names of all custom properties (cannot modify collection while iterating)
            List<string> propertyNames = new List<string>();
            foreach (DocumentProperty prop in customProperties)
            {
                propertyNames.Add(prop.Name);
            }

            // Remove each custom property by name
            foreach (string name in propertyNames)
            {
                customProperties.Remove(name);
            }

            // Save the workbook without the custom properties
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RemoveAllCustomDocumentProperties.Run();
        }
    }
}