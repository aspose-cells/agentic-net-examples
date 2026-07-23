// Title: Aspose.Cells .NET – Remove Temp_ custom properties from workbook and worksheets
// Description: C# code that loads an Excel file with Aspose.Cells, deletes all custom document and worksheet properties whose names start with "Temp_", and saves the cleaned workbook.
// Keywords: Aspose.Cells remove custom properties | delete Temp_ properties Excel .NET | clean workbook metadata C# | Aspose.Cells custom document properties | worksheet custom properties removal
// Common Searches: how to delete custom document properties starting with Temp_ using Aspose.Cells | remove worksheet custom properties by prefix C# | clean temporary Excel metadata Aspose.Cells | Aspose.Cells filter custom properties by name
// Developer Intent: Strip every custom document and worksheet property that begins with the prefix "Temp_" from an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Sanitize a workbook before distribution by removing transient metadata. | Reset a template file for reuse by clearing temporary custom properties. | Automate cleanup of generated reports that embed Temp_ markers.
// AI Prompts: Write a reusable C# method for Aspose.Cells that accepts a Workbook and a prefix string, then removes matching custom document and worksheet properties. | Generate sample code that iterates over all worksheets in a workbook and deletes custom properties whose names start with a given prefix.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertyCleanup
{
    // C# code that loads an Excel file with Aspose.Cells, deletes all custom document and worksheet properties whose names start with "Temp_", and saves the cleaned workbook.
    class Program
    {
        static void Main()
        {
            // Paths to the input and output workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook using a FileStream (lifecycle rule)
            using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                Workbook workbook = new Workbook(stream);

                // ---------- Remove custom document properties starting with "Temp_" ----------
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
                    // Use the Remove method as defined in the rule set
                    docProps.Remove(name);
                }

                // ---------- Remove worksheet custom properties starting with "Temp_" ----------
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

                    // Remove from highest index to lowest to avoid shifting issues
                    for (int i = indicesToRemove.Count - 1; i >= 0; i--)
                    {
                        // Use RemoveAt (int) as defined in the rule set
                        customProps.RemoveAt(indicesToRemove[i]);
                    }
                }

                // Save the cleaned workbook (lifecycle rule)
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
        }
    }
}
