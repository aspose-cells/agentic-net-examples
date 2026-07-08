using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsBatchExample
{
    // Simple POCO used as a data source for the templates
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class BatchProcessor
    {
        public void Run()
        {
            // Paths of the template workbooks
            string[] templateFiles = { "Template1.xlsx", "Template2.xlsx" };
            // Corresponding output file names
            string[] outputFiles = { "Result1.xlsx", "Result2.xlsx" };

            // Distinct data sources for each template
            var dataSources = new List<List<Person>>
            {
                new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 }
                },
                new List<Person>
                {
                    new Person { Name = "Charlie", Age = 35 },
                    new Person { Name = "Diana", Age = 28 }
                }
            };

            // Process each template
            for (int i = 0; i < templateFiles.Length; i++)
            {
                // Load the template workbook (uses Workbook(string) constructor)
                Workbook workbook = new Workbook(templateFiles[i]);

                // Create a WorkbookDesigner and assign the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = workbook;

                // Set a distinct data source for the current workbook
                designer.SetDataSource("Person", dataSources[i]);

                // Populate the workbook by processing smart markers
                designer.Process();

                // Save the populated workbook (uses Workbook.Save(string))
                designer.Workbook.Save(outputFiles[i]);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            new BatchProcessor().Run();
        }
    }
}