using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

class RemoveTempCustomProperties
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook from the input file
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // Access the custom document properties collection
            CustomDocumentPropertyCollection customProps = (CustomDocumentPropertyCollection)workbook.Worksheets.CustomDocumentProperties;

            // Collect names of properties that start with "Temp_"
            List<string> namesToRemove = new List<string>();
            foreach (DocumentProperty prop in customProps)
            {
                if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(prop.Name);
                }
            }

            // Remove the identified properties
            foreach (string name in namesToRemove)
            {
                customProps.Remove(name);
            }

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}