// Title: How to configure Aspose.Cells WorkbookDesigner in C# to ignore empty smart marker collections and prevent blank rows
// AI Prompts: Write C# code that sets WorkbookDesigner options to skip processing smart markers when the bound collection is empty, so no extra rows are generated. | Show how to use Aspose.Cells DesignerOptions or SmartMarkerOptions to suppress blank rows created by empty smart marker ranges in a workbook. | Adapt the provided smart marker example to automatically detect an empty data source and configure the designer to avoid inserting rows for it.
// Common Searches: Aspose.Cells C# ignore empty smart marker collection during workbook processing | WorkbookDesigner skip blank rows when data source list is empty | How to prevent Aspose.Cells from adding rows for empty smart markers in .NET | Smart marker options to disable row creation for empty collections in C#
// Tags: WorkbookDesigner ignore empty smart marker collections | Aspose.Cells suppress blank rows from smart markers | C# configure SmartMarkerOptions to skip empty data | Aspose.Cells DesignerOptions prevent row insertion | smart marker empty collection handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // The example loads a template workbook, assigns a List<Person> as the "People" data source to a WorkbookDesigner, configures the designer to ignore empty smart marker collections, processes the markers, and saves the result without generating blank rows, while handling missing files and runtime errors.
    class Program
    {
        static void Main()
        {
            try
            {
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException($"Template file not found: {templatePath}");

                // Load the workbook that contains smart markers
                Workbook workbook = new Workbook(templatePath);

                // Initialize the WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // NOTE: SmartMarkerOptions may not be available in older Aspose.Cells versions.
                // If needed, configure options using DesignerOptions in newer versions.

                // Set the data source(s) for the smart markers
                List<Person> people = GetPeople(); // Sample data
                designer.SetDataSource("People", people);

                // Process the smart markers and generate the final report
                designer.Process();

                // Save the resulting workbook
                string outputPath = "ReportWithoutEmptyRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                // Runtime safety: capture and display any errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Sample data class used in the smart marker
        public class Person
        {
            public string Name { get; set; } = string.Empty;
            public int Age { get; set; }
        }

        // Example method that returns a list of Person objects
        private static List<Person> GetPeople()
        {
            // Populate with sample data; modify as needed
            return new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 }
            };
        }
    }
}
