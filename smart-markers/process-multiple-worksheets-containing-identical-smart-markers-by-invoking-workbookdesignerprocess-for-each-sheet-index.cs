using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook that contains identical smart markers on each sheet
        Workbook workbook = new Workbook("SmartMarkerTemplate.xlsx");

        // Initialize the WorkbookDesigner and assign the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Sample data source shared by all worksheets
        var persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 }
        };

        // Bind the data source to the smart marker name used in the template
        designer.SetDataSource("Person", persons);

        // Iterate through each worksheet and process its smart markers individually
        for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
        {
            // Process only the current sheet; true = preserve unrecognized markers
            designer.Process(sheetIndex, true);
        }

        // Save the processed workbook
        workbook.Save("ProcessedOutput.xlsx");
    }

    // Simple data class used as the data source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}