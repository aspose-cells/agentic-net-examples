using System;
using System.Collections.Generic;
using Aspose.Cells;

class SmartMarkerExport
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook("template.xlsx");

        // Initialize the designer with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Prepare a sample data source
        List<Person> persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 28 }
        };

        // Bind the data source to a smart‑marker name
        designer.SetDataSource("Persons", persons);

        // Process the smart markers; preserve any unrecognized markers
        designer.Process(true);

        // Keep formulas editable by disabling automatic calculation on save
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Save the processed workbook as XLSX; formulas remain editable
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }

    // Simple POCO class used as a data source for the smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}