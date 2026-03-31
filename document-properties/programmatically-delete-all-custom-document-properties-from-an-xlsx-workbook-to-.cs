using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsMetadataCleanup
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_cleaned.xlsx";

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of custom document properties
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Collect the names of all custom properties (cannot modify collection while iterating)
            List<string> propertyNames = new List<string>();
            foreach (DocumentProperty prop in customProps)
            {
                propertyNames.Add(prop.Name);
            }

            // Remove each custom property by name
            foreach (string name in propertyNames)
            {
                customProps.Remove(name);
            }

            // Optionally remove personal information as well
            // workbook.RemovePersonalInformation();

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"All custom document properties have been removed and saved to '{outputPath}'.");
        }
    }
}