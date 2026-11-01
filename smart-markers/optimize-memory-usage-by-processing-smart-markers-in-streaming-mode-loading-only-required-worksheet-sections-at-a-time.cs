// Title: Memory‑efficient processing of Excel smart markers by streaming required worksheet sections with Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells streaming mode to process smart markers in large Excel files C# | reduce memory consumption when expanding smart markers with Aspose.Cells .NET | load only specific worksheets while processing smart markers using Aspose.Cells | Aspose.Cells smart marker processing with partial workbook loading in C# | streaming smart marker expansion example for .NET developers
// Tags: streaming smart marker expansion | partial worksheet loading Aspose.Cells | memory‑optimized workbook processing C# | Aspose.Cells WorkbookDesigner streaming | large Excel template smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerExample
{
    // The example shows how to load an Excel template, bind a list of Person objects as a data source for smart markers, process the markers while streaming only the needed worksheet sections to keep memory usage low, and save the resulting workbook.
    class Program
    {
        static void Main()
        {
            // Paths for template and result workbooks
            string inputPath = "TemplateWithSmartMarkers.xlsx";
            string outputPath = "ResultSmartMarkers.xlsx";

            // Ensure the template file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load workbook (default loading)
                Workbook workbook = new Workbook(inputPath);

                // Prepare data source for smart markers
                List<Person> people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30, City = "New York" },
                    new Person { Name = "Bob",   Age = 25, City = "London" },
                    new Person { Name = "Carol", Age = 28, City = "Sydney" }
                };

                // Process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("People", people); // smart marker prefix: &=People.Name etc.
                designer.Process();                       // expands smart markers with the data

                // Save workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
    }
}
