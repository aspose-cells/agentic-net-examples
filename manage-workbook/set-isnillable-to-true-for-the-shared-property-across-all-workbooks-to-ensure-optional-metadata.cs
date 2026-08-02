using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsSharedPropertyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in memory)
            Workbook workbook = new Workbook();

            // Add a content type property named "Shared"
            // Parameters: name, value, type (optional). Here we use a simple string value.
            int index = workbook.ContentTypeProperties.Add("Shared", "True", "Boolean");

            // Retrieve the added property
            ContentTypeProperty sharedProperty = workbook.ContentTypeProperties[index];

            // Mark the property as nillable (optional) across all workbooks
            sharedProperty.IsNillable = true;

            // Optionally, set the workbook itself as shared (not required for IsNillable)
            workbook.Settings.Shared = true;

            // Save the workbook to a file
            string outputPath = "SharedPropertyWorkbook.xlsx";
            workbook.Save(outputPath);

            // Verify the setting
            Console.WriteLine($"Content type property 'Shared' IsNillable: {sharedProperty.IsNillable}");
            Console.WriteLine($"Workbook Settings.Shared: {workbook.Settings.Shared}");
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}