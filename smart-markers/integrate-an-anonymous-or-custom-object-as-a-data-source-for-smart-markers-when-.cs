using System;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers (e.g., &=$Person.Name, &=$Person.Age)
            Workbook workbook = new Workbook("template.xlsx");

            // Create an anonymous object collection to be used as a data source
            var persons = new[]
            {
                new { Name = "John Doe", Age = 30 },
                new { Name = "Jane Smith", Age = 25 }
            };

            // Initialize the WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Bind the anonymous object collection to the smart marker variable "Person"
            designer.SetDataSource("Person", persons);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the resulting workbook
            workbook.Save("output.xlsx");
        }
    }
}