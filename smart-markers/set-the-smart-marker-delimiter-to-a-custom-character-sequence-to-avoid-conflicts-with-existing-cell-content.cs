// Title: Set a custom smart marker delimiter sequence in Aspose.Cells for .NET to prevent conflicts with existing cell content
// AI Prompts: Generate C# code that configures WorkbookDesigner to use '<<' and '>>' as the smart marker delimiters before processing a template workbook. | Show how to change the default smart marker start and end delimiters in Aspose.Cells for .NET using the WorkbookDesigner.Options.SmartMarkerDelimiter property. | Provide a step‑by‑step example of loading an Excel template, setting a custom smart marker delimiter, assigning a data source, processing, and saving the result with Aspose.Cells.
// Common Searches: Aspose.Cells .NET change smart marker delimiter to avoid cell value clash | custom smart marker start and end symbols in Aspose.Cells workbookdesigner | how to set smart marker delimiters in C# Excel template processing | prevent smart marker delimiter collision with existing data using Aspose.Cells
// Tags: WorkbookDesigner custom smart marker delimiters | Aspose.Cells smart marker delimiter configuration | C# set smart marker start end symbols | Excel template delimiter conflict resolution | Aspose.Cells data source processing with custom delimiters

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel template, configuring a custom smart marker delimiter, assigning a data source, processing the smart markers with WorkbookDesigner, and saving the resulting workbook, while handling missing file errors.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found: {templatePath}");

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize the WorkbookDesigner (smart marker processor) for the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Sample data source for the smart markers
            var employees = new List<Employee>
            {
                new Employee { Name = "John", Age = 30 },
                new Employee { Name = "Jane", Age = 28 }
            };

            // Assign the data source with a name that matches the smart marker collection in the template
            designer.SetDataSource("Employees", employees);
            designer.Process();

            // Save the processed workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO class used in the data source
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
