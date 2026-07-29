// Title: Load an Excel Template (File or MemoryStream) for Smart Markers with Aspose.Cells for .NET
// Description: Demonstrates how to verify a template file, optionally set LoadOptions, load an Excel workbook from a file path or a MemoryStream, create a WorkbookDesigner, bind a List<Person> to the "Person" smart marker, process the markers, and save the populated result as a new XLSX file using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | load workbook from file | load workbook from stream | MemoryStream Excel | smart markers | WorkbookDesigner | LoadOptions | Excel template loading | prevent file lock | bind data source | process smart markers
// Common Searches: How to load an Excel template from a MemoryStream for smart markers in C# | Aspose.Cells load workbook with LoadOptions before using smart markers | Load Excel file into WorkbookDesigner without locking the file | C# example for loading Excel template and processing smart markers | Aspose.Cells smart markers load from stream vs file
// Developer Intent: Load an Excel template (file or stream) and prepare it for smart‑marker processing in C#.
// Use Cases: Read a template file, bind a List<Person> to the &Person smart marker, process the markers, and save the output as Result.xlsx. | Load the template into a MemoryStream to avoid file locks, reset the stream position, run WorkbookDesigner, and generate the final workbook. | Apply LoadOptions (e.g., KeepUnparsedData = false) to improve loading performance before initializing WorkbookDesigner.
// AI Prompts: Show me C# code to load an Excel template from a byte array and process smart markers with Aspose.Cells. | Provide robust error handling for missing template files and correct stream positioning when using WorkbookDesigner with smart markers. | Explain how LoadOptions affect performance when loading large Excel templates for smart marker processing.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Demonstrates how to verify a template file, optionally set LoadOptions, load an Excel workbook from a file path or a MemoryStream, create a WorkbookDesigner, bind a List<Person> to the "Person" smart marker, process the markers, and save the populated result as a new XLSX file using Aspose.Cells for C#.
public class SmartMarkerExample
{
    public static void Run()
    {
        // Path to the Excel template that contains smart markers
        string templatePath = "Template.xlsx";

        // Verify that the template file exists to avoid FileNotFoundException
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        try
        {
            // Create LoadOptions – optional performance tweak for template loading
            LoadOptions loadOptions = new LoadOptions
            {
                KeepUnparsedData = false // we do not need to keep unparsed data
            };

            // ---------- Load workbook from a file ----------
            Workbook wbFromFile = new Workbook(templatePath, loadOptions);

            // ---------- Load workbook from a memory stream ----------
            Workbook wbFromStream;
            using (MemoryStream templateStream = new MemoryStream())
            {
                using (FileStream fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    fs.CopyTo(templateStream);
                }
                templateStream.Position = 0; // reset stream position for reading
                wbFromStream = new Workbook(templateStream, loadOptions);
            }

            // Choose which workbook to work with (file‑loaded or stream‑loaded)
            Workbook workbook = wbFromFile; // or wbFromStream

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Sample data source that matches the smart marker name in the template (e.g., &Person.Name, &Person.Age)
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 }
            };

            // Bind the data source to the smart marker name "Person"
            designer.SetDataSource("Person", persons);

            // Process the smart markers – this populates the worksheet with data
            designer.Process();

            // Save the processed workbook to a new file
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}

// Entry point for the application
public class Program
{
    public static void Main()
    {
        SmartMarkerExample.Run();
    }
}
