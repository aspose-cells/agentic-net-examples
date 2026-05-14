using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load the template workbook (built‑in properties are read from this file)
        Workbook templateWorkbook = new Workbook("Template.xlsx");

        // Create a new empty workbook
        Workbook newWorkbook = new Workbook();

        // Clone all built‑in document properties from the template to the new workbook
        foreach (DocumentProperty sourceProp in templateWorkbook.BuiltInDocumentProperties)
        {
            // The destination workbook has the same set of built‑in property names,
            // so we can assign the value directly.
            newWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
        }

        // Save the new workbook with the cloned properties
        newWorkbook.Save("ClonedProperties.xlsx");
    }
}