// Title: Export a workbook processed with Aspose.Cells smart markers to XLSX while keeping formulas editable in C#
// AI Prompts: Generate C# code that loads an XLSX template, binds a collection to smart markers via WorkbookDesigner, turns off automatic formula evaluation, and saves the workbook as a new XLSX file. | Demonstrate how to preserve editable formulas after processing smart markers with Aspose.Cells by setting CalculateOnSave to false before saving.
// Common Searches: C# Aspose.Cells keep formulas editable when saving workbook after smart marker processing | How to disable CalculateOnSave in Aspose.Cells after using WorkbookDesigner | Export smart marker populated workbook to XLSX without evaluating formulas Aspose.Cells .NET | Aspose.Cells preserve formula cells when saving processed smart markers
// Tags: process smart markers Aspose.Cells C# | save workbook as xlsx Aspose.Cells | disable formula calculation on save Aspose.Cells | WorkbookDesigner data binding C# | keep formulas editable Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

// // Loads a template workbook, binds a list of Person objects to smart markers using WorkbookDesigner, processes the markers, disables automatic formula calculation on save, and saves the result as an XLSX file with formulas remaining editable.
class SmartMarkerExport
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Create a WorkbookDesigner and assign the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Prepare a simple data source (list of Person objects)
        ArrayList persons = new ArrayList();
        persons.Add(new Person { Name = "John Doe", Age = 30 });
        persons.Add(new Person { Name = "Jane Smith", Age = 28 });

        // Bind the data source to the name used in the smart markers
        designer.SetDataSource("Persons", persons);

        // Process all smart markers in the workbook
        designer.Process();

        // Keep formulas editable (do not force calculation on save)
        workbook.Settings.FormulaSettings.CalculateOnSave = false;

        // Save the processed workbook to XLSX format
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }

    // POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
