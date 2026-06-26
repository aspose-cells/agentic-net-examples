using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class CloneBuiltInProperties
{
    static void Main()
    {
        // Load the template workbook (source)
        Workbook sourceWorkbook = new Workbook("TemplateWorkbook.xlsx");

        // Create a new empty workbook (destination)
        Workbook destWorkbook = new Workbook();

        // Clone each built‑in document property from source to destination
        foreach (DocumentProperty sourceProp in sourceWorkbook.BuiltInDocumentProperties)
        {
            // Ensure the destination has the same property (built‑in collection always contains the same set)
            DocumentProperty destProp = destWorkbook.BuiltInDocumentProperties[sourceProp.Name];
            destProp.Value = sourceProp.Value;
        }

        // Save the destination workbook with the cloned properties
        destWorkbook.Save("ClonedPropertiesWorkbook.xlsx");
    }
}