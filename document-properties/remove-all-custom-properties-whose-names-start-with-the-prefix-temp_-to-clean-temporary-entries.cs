// Title: Aspose.Cells for .NET – Delete Temp_ custom properties from Excel workbooks
// Description: C# sample that loads an XLSX file, removes all workbook‑level and worksheet‑level custom properties whose names start with "Temp_", and saves the cleaned file using Aspose.Cells.
// Keywords: Aspose.Cells remove custom document property | delete temporary Excel property .NET | clean workbook custom properties | C# Aspose.Cells remove prefix Temp_ | Excel metadata cleanup Aspose
// Common Searches: how to delete custom document properties with a prefix in Aspose.Cells | remove worksheet custom properties Temp_ C# | Aspose.Cells clean temporary metadata from Excel | C# code to purge Temp_ properties from workbook
// Developer Intent: Programmatically strip any custom document or worksheet property that begins with "Temp_" and persist the updated workbook.
// Use Cases: Sanitize generated reports before distribution to avoid leaking internal identifiers. | Batch‑process a library of spreadsheets to eliminate placeholder metadata. | Prepare Excel files for compliance audits by removing transient properties.
// AI Prompts: Generate C# code using Aspose.Cells to remove all custom properties prefixed with "Temp_" from a workbook and its sheets. | Explain safe iteration techniques for deleting items from DocumentPropertyCollection and Worksheet.CustomProperties. | Show how to log the names of removed properties while cleaning an Excel file with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

// C# sample that loads an XLSX file, removes all workbook‑level and worksheet‑level custom properties whose names start with "Temp_", and saves the cleaned file using Aspose.Cells.
class RemoveTempCustomProperties
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook (lifecycle rule: load)
        using (FileStream stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            Workbook workbook = new Workbook(stream);

            // Access the custom document properties collection
            DocumentPropertyCollection customProps = workbook.Worksheets.CustomDocumentProperties;

            // Collect names of properties that start with "Temp_"
            List<string> namesToRemove = new List<string>();
            foreach (DocumentProperty prop in customProps)
            {
                if (prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(prop.Name);
                }
            }

            // Remove the identified properties using DocumentPropertyCollection.Remove
            foreach (string name in namesToRemove)
            {
                customProps.Remove(name);
            }

            // Optionally, also clean worksheet-level custom properties
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                CustomPropertyCollection sheetProps = sheet.CustomProperties;
                // Gather indices to remove (iterate backwards to avoid index shift)
                for (int i = sheetProps.Count - 1; i >= 0; i--)
                {
                    if (sheetProps[i].Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        sheetProps.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        Console.WriteLine("Temporary custom properties removed and workbook saved to: " + outputPath);
    }
}
