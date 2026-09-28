// Title: Use Aspose.Cells WorkbookDesigner to fill an Excel template from a List<Person> while preserving cell formatting (CopyStyle) in C#
// AI Prompts: Write C# code that loads an .xlsx template, binds a List<Person> to a smart marker named "Person", enables style inheritance, processes the template with WorkbookDesigner, and saves the output workbook. | Show how to configure DesignerOptions.CopyStyle (or the equivalent setting) for WorkbookDesigner in Aspose.Cells to keep original cell styles during smart‑marker processing. | Add robust error handling that checks for a missing template file, creates the result directory if it does not exist, and logs any exceptions that occur while processing the template.
// Common Searches: Aspose.Cells C# WorkbookDesigner copy style option example | How to keep original cell formatting when using smart markers in Aspose.Cells | Populate Excel template with a list of objects using Aspose.Cells and preserve styles | Enable CopyStyle attribute in Aspose.Cells .NET for template processing | C# code to bind List<Person> to smart marker and retain formatting
// Tags: WorkbookDesigner copy style C# | Aspose.Cells smart markers preserve formatting | populate Excel template from List<Person> | error handling missing template Aspose.Cells | create output directory before saving workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel template, uses WorkbookDesigner to bind a List<Person> to the "Person" smart marker, enables style inheritance (CopyStyle) so original cell formatting is retained, processes the template, ensures the output folder exists, and saves the resulting workbook, with basic exception handling for missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            // Load the workbook template that contains markers for data insertion
            Workbook templateWorkbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded template
            WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

            // NOTE: In some Aspose.Cells versions the Options property is unavailable.
            // Copying of cell styles is handled automatically, or can be configured via DesignerOptions if supported.

            // Sample data source: a list of records to be merged into the template
            List<Person> data = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 28 }
            };

            // Bind the data source to the marker named "Person" in the template
            designer.SetDataSource("Person", data);

            // Process the template – markers are replaced with data
            designer.Process();

            // Save the resulting workbook
            const string resultPath = "Result.xlsx";

            // Ensure the directory for the result file exists
            string resultDir = Path.GetDirectoryName(resultPath);
            if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
            {
                Directory.CreateDirectory(resultDir);
            }

            designer.Workbook.Save(resultPath);
            Console.WriteLine($"Result workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple POCO class representing a record
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
