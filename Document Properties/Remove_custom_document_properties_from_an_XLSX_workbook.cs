using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class RemoveCustomDocumentProperties
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook from a file stream
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // Access the custom document properties collection
            CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

            // Gather all property names (cannot modify collection while iterating)
            List<string> namesToRemove = new List<string>();
            foreach (DocumentProperty prop in customProps)
            {
                namesToRemove.Add(prop.Name);
            }

            // Remove each custom property by name
            foreach (string name in namesToRemove)
            {
                customProps.Remove(name);
            }

            // Save the workbook with the properties removed
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}