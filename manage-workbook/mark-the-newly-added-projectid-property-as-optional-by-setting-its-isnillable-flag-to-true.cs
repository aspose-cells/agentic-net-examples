using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsProjectIdExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Add the newly introduced ProjectId content type property
            // Parameters: name, value, type (optional, here using "string")
            workbook.ContentTypeProperties.Add("ProjectId", "12345", "string");

            // Retrieve the added property to modify its attributes
            ContentTypeProperty projectIdProperty = workbook.ContentTypeProperties["ProjectId"];

            // Mark the property as optional (allow empty/nil values)
            projectIdProperty.IsNillable = true;

            // Save the workbook to verify the property settings
            workbook.Save("ProjectIdOptional.xlsx");

            // Optional: output confirmation
            Console.WriteLine($"Property 'ProjectId' IsNillable set to: {projectIdProperty.IsNillable}");
        }
    }
}