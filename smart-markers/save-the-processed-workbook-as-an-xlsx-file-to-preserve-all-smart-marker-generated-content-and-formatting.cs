using System;
using System.Collections;
using Aspose.Cells;

namespace SmartMarkerSaveExample
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains smart markers (template.xlsx should exist in the execution folder)
            Workbook workbook = new Workbook("template.xlsx");

            // Create a data source – for demonstration we use an ArrayList of simple objects
            ArrayList persons = new ArrayList();

            // Sample data class
            var person1 = new { Name = "John Doe", Age = 30 };
            var person2 = new { Name = "Jane Smith", Age = 28 };
            persons.Add(person1);
            persons.Add(person2);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the data source for the smart markers (the name "Persons" must match the marker prefix in the template)
            designer.SetDataSource("Persons", persons);

            // Process all smart markers in the workbook
            designer.Process();

            // Save the processed workbook as XLSX to preserve all generated content and formatting
            // This uses the provided Save(string, SaveFormat) method as required by the rules
            workbook.Save("ProcessedOutput.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved successfully as ProcessedOutput.xlsx");
        }
    }
}