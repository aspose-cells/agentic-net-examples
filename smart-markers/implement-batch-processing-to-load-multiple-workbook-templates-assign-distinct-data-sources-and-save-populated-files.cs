// Title: Batch generate Excel files from multiple templates using Aspose.Cells Smart Markers (C#/.NET)
// Description: C# code that loads each template workbook, binds a unique List<Person> to the "Person" smart marker via WorkbookDesigner, processes the markers, and saves the populated file while handling missing files and runtime exceptions.
// Keywords: Aspose.Cells | C# | .NET | smart markers | WorkbookDesigner | batch processing | multiple templates | populate Excel | data source binding | Excel automation
// Common Searches: aspocells batch smart markers c# | load multiple excel templates aspnet | bind different data sources to each workbook | process several excel files with smart markers | automate excel report generation aspocells
// Developer Intent: Build a routine that fills several Excel templates with distinct data sets using smart markers.
// Use Cases: Generate department‑specific reports by applying each department’s employee list to its own template. | Create client‑wise invoices where each client has a custom layout and a unique order collection. | Produce a series of certificates with varied designs, binding each design to its list of recipient details.
// AI Prompts: Refactor the batch processor to use a dictionary that maps template paths to output paths and a matching dictionary for data sources. | Add logging that writes the processing status of each workbook to a CSV file while ensuring the loop continues after errors. | Explain how to parallelize the workbook processing using async/await or Task Parallel Library with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // C# code that loads each template workbook, binds a unique List<Person> to the "Person" smart marker via WorkbookDesigner, processes the markers, and saves the populated file while handling missing files and runtime exceptions.
    public class BatchWorkbookProcessor
    {
        public static void Run()
        {
            // Paths to template workbooks and the desired output files
            string[] templateFiles = { "Template1.xlsx", "Template2.xlsx" };
            string[] outputFiles   = { "Result1.xlsx",   "Result2.xlsx"   };

            // Distinct data sources for each template (using the same marker name "Person")
            var dataSources = new List<object>
            {
                new List<Person>
                {
                    new Person("Alice", 28),
                    new Person("Bob",   35)
                },
                new List<Person>
                {
                    new Person("Charlie", 22),
                    new Person("Diana",   31)
                }
            };

            // Process each template independently
            for (int i = 0; i < templateFiles.Length; i++)
            {
                try
                {
                    // Verify template file exists
                    if (!File.Exists(templateFiles[i]))
                    {
                        Console.WriteLine($"Template file not found: {templateFiles[i]}");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templateFiles[i]);

                    // Create a designer bound to the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);

                    // Bind the corresponding data source
                    designer.SetDataSource("Person", dataSources[i]);

                    // Populate the smart markers
                    designer.Process();

                    // Save the populated workbook
                    designer.Workbook.Save(outputFiles[i]);

                    Console.WriteLine($"Successfully generated: {outputFiles[i]}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing {templateFiles[i]}: {ex.Message}");
                }
            }
        }
    }

    // Simple POCO used as a data source for demonstration
    public class Person
    {
        public string Name { get; set; }
        public int    Age  { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age  = age;
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                BatchWorkbookProcessor.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
