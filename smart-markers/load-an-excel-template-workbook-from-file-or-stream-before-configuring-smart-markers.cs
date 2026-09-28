// Title: Load an Excel template workbook from a file or MemoryStream and process smart markers with Aspose.Cells in C#
// AI Prompts: Load a template workbook from a file path, bind a List<Person> to a smart marker using WorkbookDesigner, process the markers, and save the result as an XLSX file. | Read the same Excel template into a MemoryStream, create a Workbook from the stream, apply smart‑marker processing with a POCO data source, and write the output to a separate file. | Add robust error handling for missing template files and load failures when using both file‑based and stream‑based constructors with Aspose.Cells.
// Common Searches: Aspose.Cells load Excel template from MemoryStream for smart markers C# | how to bind a POCO list to a smart marker using WorkbookDesigner in C# | processing smart markers in a workbook loaded from a file with Aspose.Cells | error handling when loading Excel workbook from stream in Aspose.Cells | save processed smart marker workbook to a new XLSX file using Aspose.Cells C#
// Tags: file path workbook loading Aspose.Cells C# | memory stream workbook loading Aspose.Cells | WorkbookDesigner smart marker processing | POCO list binding to smart marker | export processed workbook to xlsx

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example checks that the Excel template exists, loads it either directly from a file path or via a MemoryStream into an Aspose.Cells Workbook, binds a List<Person> to a smart‑marker named "Person" with WorkbookDesigner, processes all smart markers, and saves each processed workbook to separate XLSX files, demonstrating proper error handling for missing files and load failures.
public class SmartMarkerExample
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string templatePath = "TemplateSmartMarkers.xlsx";

        // Verify that the template file exists before attempting to load it.
        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file \"{templatePath}\" not found.");
            return;
        }

        // Load the template workbook from a file path.
        Workbook workbookFromFile;
        try
        {
            workbookFromFile = new Workbook(templatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook from file: {ex.Message}");
            return;
        }

        ProcessSmartMarkers(workbookFromFile);
        // Save the processed workbook.
        workbookFromFile.Save("ResultFromFile.xlsx", SaveFormat.Xlsx);

        // Load the same template workbook from a memory stream.
        using (MemoryStream stream = new MemoryStream())
        {
            // Copy the template file into the memory stream.
            try
            {
                using (FileStream fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(stream);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read template into stream: {ex.Message}");
                return;
            }

            stream.Position = 0; // Reset stream position for reading.

            // Create a Workbook instance using the Stream constructor.
            Workbook workbookFromStream;
            try
            {
                workbookFromStream = new Workbook(stream);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook from stream: {ex.Message}");
                return;
            }

            ProcessSmartMarkers(workbookFromStream);
            // Save the processed workbook.
            workbookFromStream.Save("ResultFromStream.xlsx", SaveFormat.Xlsx);
        }
    }

    private static void ProcessSmartMarkers(Workbook workbook)
    {
        // Initialize WorkbookDesigner and assign the loaded workbook.
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };

        // Sample data source for smart markers.
        List<Person> persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 28 }
        };

        // Bind the data source to the marker name "Person".
        designer.SetDataSource("Person", persons);

        // Process all smart markers in the workbook.
        designer.Process();
    }

    // Simple POCO class used as a data source for smart markers.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
