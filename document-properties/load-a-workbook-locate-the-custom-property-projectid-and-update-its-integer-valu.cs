using Aspose.Cells;

class UpdateCustomProperty
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of custom document properties
        var customProps = workbook.CustomDocumentProperties;

        // Locate the custom property named "ProjectId"
        var projectIdProp = customProps["ProjectId"];

        // Update its value or add it if it does not exist
        if (projectIdProp != null)
        {
            projectIdProp.Value = 12345;
        }
        else
        {
            customProps.Add("ProjectId", 12345);
        }

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}