using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyCleaner
{
    class Program
    {
        static void Main()
        {
            // Paths for input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook from the input file
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                Workbook workbook = new Workbook(stream);

                // ----- Remove custom document properties starting with "Temp_" -----
                DocumentPropertyCollection docProps = workbook.Worksheets.CustomDocumentProperties;
                List<string> docPropNamesToRemove = new List<string>();

                foreach (DocumentProperty prop in docProps)
                {
                    if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        docPropNamesToRemove.Add(prop.Name);
                    }
                }

                foreach (string name in docPropNamesToRemove)
                {
                    docProps.Remove(name);
                }

                // ----- Remove worksheet custom properties starting with "Temp_" -----
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CustomPropertyCollection customProps = sheet.CustomProperties;
                    List<int> indicesToRemove = new List<int>();

                    for (int i = 0; i < customProps.Count; i++)
                    {
                        if (customProps[i].Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                        {
                            indicesToRemove.Add(i);
                        }
                    }

                    // Remove by index in descending order to keep remaining indices valid
                    for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                    {
                        customProps.RemoveAt(indicesToRemove[i]);
                    }
                }

                // Save the cleaned workbook to the output file
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
        }
    }
}