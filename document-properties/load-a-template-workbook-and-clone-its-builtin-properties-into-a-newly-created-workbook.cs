using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCloneBuiltInProperties
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook (source)
            Workbook sourceWorkbook = new Workbook("TemplateWorkbook.xlsx");

            // Create a new empty workbook (destination)
            Workbook destWorkbook = new Workbook();

            // Clone built‑in document properties from source to destination
            foreach (DocumentProperty sourceProp in sourceWorkbook.BuiltInDocumentProperties)
            {
                // Ensure the destination has the same property and assign its value
                destWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
            }

            // Save the destination workbook with the cloned properties
            destWorkbook.Save("ClonedPropertiesWorkbook.xlsx");
        }
    }
}