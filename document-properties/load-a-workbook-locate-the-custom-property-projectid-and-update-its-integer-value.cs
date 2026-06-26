using Aspose.Cells;
using Aspose.Cells.Properties;

class UpdateProjectId
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Locate the property named "ProjectId"
        DocumentProperty projectIdProp = customProps["ProjectId"];
        if (projectIdProp != null)
        {
            // Update its integer value
            projectIdProp.Value = 12345; // new integer value
        }

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}