using System;
using System.Collections;
using Aspose.Cells;

namespace SmartMarkerSaveExample
{
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare a simple data source (could be any collection, DataTable, JSON, etc.)
            ArrayList persons = new ArrayList();
            persons.Add(new Person { Name = "John Doe", Age = 30 });
            persons.Add(new Person { Name = "Jane Smith", Age = 28 });

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Assign the data source to a name used in the smart markers
            designer.SetDataSource("Persons", persons);

            // Process the smart markers – this populates the worksheet with data
            designer.Process();

            // Save the processed workbook as XLSX to preserve all generated content and formatting
            // Using the Save(string, SaveFormat) overload as defined in the Aspose.Cells API
            workbook.Save("ProcessedResult.xlsx", SaveFormat.Xlsx);
        }
    }

    // Simple POCO class used as a data source for the smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}