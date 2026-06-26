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

        // Load the workbook from a file stream
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // ---------- Remove Custom Document Properties ----------
            // Collect names that start with "Temp_"
            DocumentPropertyCollection docProps = workbook.Worksheets.CustomDocumentProperties;
            List<string> docPropNamesToRemove = new List<string>();
            foreach (DocumentProperty prop in docProps)
            {
                if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    docPropNamesToRemove.Add(prop.Name);
                }
            }
            // Remove the collected properties
            foreach (string name in docPropNamesToRemove)
            {
                docProps.Remove(name);
            }

            // ---------- Remove Worksheet Custom Properties ----------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                CustomPropertyCollection customProps = sheet.CustomProperties;
                List<int> indicesToRemove = new List<int>();

                // Identify indices of properties whose names start with "Temp_"
                for (int i = 0; i < customProps.Count; i++)
                {
                    if (customProps[i].Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        indicesToRemove.Add(i);
                    }
                }

                // Remove from highest index to lowest to avoid shifting issues
                for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                {
                    customProps.RemoveAt(indicesToRemove[i]);
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}