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

            // Add the new ProjectId content type property (value can be any placeholder)
            // Overload with name, value, and type (optional)
            workbook.ContentTypeProperties.Add("ProjectId", "12345", "string");

            // Retrieve the added property
            ContentTypeProperty projectIdProperty = workbook.ContentTypeProperties["ProjectId"];

            // Mark the property as optional (allow empty values)
            projectIdProperty.IsNillable = true;

            // Optional: verify the flag
            Console.WriteLine($"ProjectId IsNillable: {projectIdProperty.IsNillable}");

            // Save the workbook to a file
            workbook.Save("ProjectIdOptionalDemo.xlsx");
        }
    }
}