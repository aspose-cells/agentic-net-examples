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

            // Add the new ProjectId content type property (value and type can be set as needed)
            // Here we add it with a placeholder value and type "string"
            workbook.ContentTypeProperties.Add("ProjectId", "12345", "string");

            // Retrieve the added property
            ContentTypeProperty projectIdProperty = workbook.ContentTypeProperties["ProjectId"];

            // Mark the property as optional by setting IsNillable to true
            projectIdProperty.IsNillable = true;

            // Optional: output the flag to verify
            Console.WriteLine($"ProjectId IsNillable: {projectIdProperty.IsNillable}");

            // Save the workbook to a file
            workbook.Save("ProjectIdOptionalDemo.xlsx");
        }
    }
}