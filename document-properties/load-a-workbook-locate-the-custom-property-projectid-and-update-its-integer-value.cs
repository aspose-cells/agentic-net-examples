using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class UpdateProjectIdProperty
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Try to retrieve the "ProjectId" property
        DocumentProperty projectIdProp = customProps["ProjectId"];

        if (projectIdProp != null)
        {
            // Property exists – update its integer value
            projectIdProp.Value = 12345;   // new integer value
        }
        else
        {
            // Property does not exist – add it with the desired integer value
            customProps.Add("ProjectId", 12345);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}