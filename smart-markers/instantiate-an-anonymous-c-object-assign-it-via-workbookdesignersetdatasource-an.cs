using System;
using Aspose.Cells;
using System.Collections.Generic;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the XLSX template that contains smart markers (e.g., &Person.Name, &Person.Age)
            Workbook templateWorkbook = new Workbook("Template.xlsx");

            // Create a WorkbookDesigner and assign the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = templateWorkbook;

            // Prepare an anonymous object collection as the data source
            var persons = new[]
            {
                new { Name = "John Doe", Age = 30 },
                new { Name = "Jane Smith", Age = 25 }
            };

            // Bind the anonymous collection to the smart marker variable "Person"
            designer.SetDataSource("Person", persons);

            // Process the smart markers and populate the worksheet with data
            designer.Process();

            // Save the populated workbook
            designer.Workbook.Save("Output.xlsx");
        }
    }
}