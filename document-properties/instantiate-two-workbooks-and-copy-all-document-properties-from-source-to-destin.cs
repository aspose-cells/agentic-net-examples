using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load the source workbook from a file
        Workbook sourceWorkbook = new Workbook("source.xlsx");

        // Create an empty destination workbook
        Workbook destinationWorkbook = new Workbook();

        // ----- Copy Built‑in Document Properties -----
        foreach (DocumentProperty sourceProp in sourceWorkbook.BuiltInDocumentProperties)
        {
            destinationWorkbook.BuiltInDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
        }

        // ----- Copy Custom Document Properties -----
        foreach (DocumentProperty sourceProp in sourceWorkbook.CustomDocumentProperties)
        {
            if (destinationWorkbook.CustomDocumentProperties.Contains(sourceProp.Name))
            {
                destinationWorkbook.CustomDocumentProperties[sourceProp.Name].Value = sourceProp.Value;
            }
            else
            {
                // Add new custom property (value must be a string)
                destinationWorkbook.CustomDocumentProperties.Add(sourceProp.Name, sourceProp.Value?.ToString() ?? string.Empty);
            }
        }

        // Save the destination workbook with all copied properties
        destinationWorkbook.Save("destination.xlsx");
    }
}