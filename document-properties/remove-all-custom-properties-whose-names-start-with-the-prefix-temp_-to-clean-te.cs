using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyCleanup
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_cleaned.xlsx";

            // Load the workbook from a file stream
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(stream);

                // ----- Remove custom document properties with prefix "Temp_" -----
                DocumentPropertyCollection docProps = workbook.Worksheets.CustomDocumentProperties;

                // Collect names to remove (cannot modify collection while iterating)
                var namesToRemove = new System.Collections.Generic.List<string>();
                foreach (DocumentProperty prop in docProps)
                {
                    if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        namesToRemove.Add(prop.Name);
                    }
                }

                // Remove each collected property by name
                foreach (string name in namesToRemove)
                {
                    docProps.Remove(name);
                }

                // ----- Remove worksheet custom properties with prefix "Temp_" -----
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CustomPropertyCollection customProps = sheet.CustomProperties;

                    // Iterate backwards to safely remove by index
                    for (int i = customProps.Count - 1; i >= 0; i--)
                    {
                        CustomProperty prop = customProps[i];
                        if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                        {
                            customProps.RemoveAt(i);
                        }
                    }
                }

                // Save the cleaned workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Custom properties with prefix \"Temp_\" have been removed.");
        }
    }
}