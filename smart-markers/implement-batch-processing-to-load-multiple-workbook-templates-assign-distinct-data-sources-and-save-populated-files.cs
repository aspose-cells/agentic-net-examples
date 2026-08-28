// Title: Batch generate Excel files from multiple templates with individual smart‑marker data sources using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a list of template workbook paths, assigns a separate List<Person> as the 'Employees' smart‑marker data source for each workbook using WorkbookDesigner, processes the markers, and saves the result to corresponding output files. | Create a robust Aspose.Cells batch routine that checks each template file exists, loads it into a Workbook, sets a distinct data source, calls Designer.Process(), and implements try‑catch logging for any failures. | Generate a reusable method that accepts arrays of template filenames, output filenames, and data source objects, then performs smart‑marker processing for each entry with Aspose.Cells.
// Common Searches: asp.net c# generate reports from several Excel templates using Aspose.Cells smart markers | assign unique data collections to each Excel template when using Aspose.Cells | loop through multiple workbook templates and populate them with different data sets in C# | how to verify Excel template files exist before processing with Aspose.Cells | error handling strategy for batch Excel generation with Aspose.Cells
// Tags: batch processing smart markers Aspose.Cells | assign distinct data source per workbook | process multiple Excel templates C# | validate template file existence Aspose.Cells | exception handling Aspose.Cells batch generation

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcessing
{
    // Sample data class used as a data source for the smart markers
    // The example defines a Person class and three List<Person> collections, then iterates over parallel arrays of template paths, output paths, and data sources. For each entry it verifies the template file exists, loads it into a Workbook, uses WorkbookDesigner to bind the 'Employees' smart‑marker data source, processes the markers, saves the populated workbook, and logs progress while handling exceptions.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }

        public Person(string name, int age, string department)
        {
            Name = name;
            Age = age;
            Department = department;
        }
    }

    public class BatchProcessor
    {
        public static void Run()
        {
            // Define the template files, output files and the corresponding data sources
            string[] templateFiles = { "Template1.xlsx", "Template2.xlsx", "Template3.xlsx" };
            string[] outputFiles   = { "Result1.xlsx",   "Result2.xlsx",   "Result3.xlsx"   };
            string   dataSourceName = "Employees";

            // Prepare distinct data sources for each template
            var dataSource1 = new List<Person>
            {
                new Person("John Doe", 30, "Sales"),
                new Person("Jane Smith", 28, "Marketing")
            };

            var dataSource2 = new List<Person>
            {
                new Person("Alice Brown", 35, "HR"),
                new Person("Bob Johnson", 40, "Finance"),
                new Person("Carol White", 25, "IT")
            };

            var dataSource3 = new List<Person>
            {
                new Person("David Green", 45, "Operations")
            };

            object[] dataSources = { dataSource1, dataSource2, dataSource3 };

            // Ensure the arrays have the same length
            if (templateFiles.Length != outputFiles.Length || templateFiles.Length != dataSources.Length)
                throw new InvalidOperationException("Template, output and data source arrays must have the same length.");

            // Process each workbook template
            for (int i = 0; i < templateFiles.Length; i++)
            {
                try
                {
                    // Verify template file exists
                    if (!File.Exists(templateFiles[i]))
                    {
                        Console.WriteLine($"Template file '{templateFiles[i]}' not found. Skipping this entry.");
                        continue;
                    }

                    // Load the template workbook
                    Workbook workbook = new Workbook(templateFiles[i]);

                    // Initialize the WorkbookDesigner and assign the loaded workbook
                    WorkbookDesigner designer = new WorkbookDesigner
                    {
                        Workbook = workbook
                    };

                    // Set the distinct data source for this workbook
                    designer.SetDataSource(dataSourceName, dataSources[i]);

                    // Process smart markers in the workbook
                    designer.Process();

                    // Save the populated workbook to the specified output file
                    designer.Workbook.Save(outputFiles[i]);

                    Console.WriteLine($"Processed '{templateFiles[i]}' and saved as '{outputFiles[i]}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{templateFiles[i]}': {ex.Message}");
                }
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            BatchProcessor.Run();
        }
    }
}
