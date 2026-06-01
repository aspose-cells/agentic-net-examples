using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetIsNillableForSharedProperty
{
    static void Main()
    {
        // Paths of the workbooks to be processed
        string[] workbookFiles = { "Workbook1.xlsx", "Workbook2.xlsx", "Workbook3.xlsx" };

        foreach (string filePath in workbookFiles)
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(filePath);

            // Add (or overwrite) a content type property named "Shared"
            // The value and type are arbitrary; the key part is setting IsNillable
            int propIndex = workbook.ContentTypeProperties.Add("Shared", "Aspose", "text");

            // Retrieve the newly added property and mark it as nillable (optional)
            ContentTypeProperty sharedProperty = workbook.ContentTypeProperties[propIndex];
            sharedProperty.IsNillable = true;

            // Save the workbook with the updated metadata
            workbook.Save(filePath);
        }
    }
}