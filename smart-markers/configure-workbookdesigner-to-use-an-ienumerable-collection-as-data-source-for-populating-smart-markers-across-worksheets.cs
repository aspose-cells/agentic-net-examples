// Title: Use WorkbookDesigner to bind an IEnumerable<Person> list as a smart‑marker data source and populate all worksheets in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a template workbook, assigns a List<Person> to WorkbookDesigner with the name "people", processes smart markers on every sheet, and saves the populated file. | Show how to configure WorkbookDesigner with an IEnumerable collection to fill smart markers like &=people.Name and &=people.Age across multiple worksheets using Aspose.Cells.
// Common Searches: asp.net bind list of objects to WorkbookDesigner smart markers | aspose.cells populate smart markers from IEnumerable collection across all sheets | c# set data source for smart markers using WorkbookDesigner | how to use smart markers with List<T> in an Aspose.Cells template | process smart markers in multiple worksheets with WorkbookDesigner
// Tags: WorkbookDesigner set data source IEnumerable | populate smart markers from List<Person> | process smart markers across all worksheets | Aspose.Cells smart marker collection binding | C# load template and fill smart markers

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // Sample data class that will be used as the data source
    // The program creates a List<Person>, loads a workbook template containing smart markers, assigns the list as the "people" data source to a WorkbookDesigner, processes the markers on every worksheet, and saves the populated workbook.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // 1. Prepare the data source – an IEnumerable collection (List<Person> in this case)
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 },
                new Person { Name = "Carol", Age = 28 }
            };

            // 2. Load the workbook that contains smart markers.
            //    The workbook can be created beforehand with smart markers like:
            //    &=people.Name   and   &=people.Age   placed in cells.
            Workbook workbook = new Workbook("SmartMarkerTemplate.xlsx");

            // 3. Create a WorkbookDesigner and assign the workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // 4. Set the IEnumerable collection as the data source for the smart markers.
            //    The name "people" must match the smart marker prefix used in the template.
            designer.SetDataSource("people", people);

            // 5. Process the smart markers – this will populate the data across all worksheets.
            designer.Process();

            // 6. Save the populated workbook.
            designer.Workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}
