using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Sample data class
    public class Person
    {
        public string Name { get; set; }
        public bool ShowScore { get; set; }
        public double Score { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers.
            // The template should have the following cells (for example):
            // A1: "Name"
            // B1: "Score"
            // A2: "&=$Name"
            // B2: "=IF($ShowScore,$Score,\"\")"
            Workbook workbook = new Workbook("Template.xlsx");

            // Prepare data source.
            var persons = new List<Person>
            {
                new Person { Name = "Alice", ShowScore = true,  Score = 85.5 },
                new Person { Name = "Bob",   ShowScore = false, Score = 72.0 },
                new Person { Name = "Carol", ShowScore = true,  Score = 93.2 }
            };

            // Set the data source for the smart markers.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Persons", persons);

            // Process the smart markers. The IF condition inside the formula will
            // display the Score only when ShowScore is true; otherwise the cell will be empty.
            designer.Process();

            // Save the result.
            workbook.Save("Result.xlsx");
        }
    }
}