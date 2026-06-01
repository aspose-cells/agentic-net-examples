using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Cells;

namespace AsposeCellsMultiThreadedDesignerDemo
{
    // Simple data class used as a data source for the smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Person(string name, int age) { Name = name; Age = age; }
    }

    class Program
    {
        // Number of templates to process concurrently
        private const int TemplateCount = 3;

        // Paths to template files (must contain smart markers like &Person.Name, &Person.Age)
        private static readonly string[] TemplateFiles = new string[]
        {
            "Template1.xlsx",
            "Template2.xlsx",
            "Template3.xlsx"
        };

        // Array to hold the processed workbooks produced by each thread
        private static readonly Workbook[] ProcessedWorkbooks = new Workbook[TemplateCount];

        static void Main()
        {
            try
            {
                // Prepare sample data source (same for all templates in this example)
                List<Person> persons = new List<Person>
                {
                    new Person("John Doe", 30),
                    new Person("Jane Smith", 28),
                    new Person("Bob Johnson", 35)
                };

                // Create and start a thread for each template
                Thread[] threads = new Thread[TemplateCount];
                for (int i = 0; i < TemplateCount; i++)
                {
                    int index = i; // capture loop variable
                    threads[i] = new Thread(() => ProcessTemplate(index, persons));
                    threads[i].Start();
                }

                // Wait for all threads to finish
                foreach (Thread t in threads)
                    t.Join();

                // Merge all processed workbooks into a single workbook
                Workbook finalWorkbook = new Workbook(); // empty workbook
                // Remove the default empty sheet (if any)
                if (finalWorkbook.Worksheets.Count > 0)
                    finalWorkbook.Worksheets.RemoveAt(0);

                for (int i = 0; i < TemplateCount; i++)
                {
                    Workbook source = ProcessedWorkbooks[i];
                    if (source == null) continue; // skip if processing failed

                    foreach (Worksheet ws in source.Worksheets)
                    {
                        // AddCopy expects the sheet name, not the Worksheet object
                        finalWorkbook.Worksheets.AddCopy(ws.Name);
                    }
                }

                // Ensure there is at least one worksheet before saving
                if (finalWorkbook.Worksheets.Count == 0)
                    finalWorkbook.Worksheets.Add("Sheet1");

                // Save the merged result
                finalWorkbook.Save("MergedResult.xlsx");
                Console.WriteLine("Merged workbook saved as MergedResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Method executed by each thread: loads a template, processes it, and stores the result
        private static void ProcessTemplate(int templateIndex, List<Person> dataSource)
        {
            try
            {
                string templatePath = TemplateFiles[templateIndex];

                // Prevent FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook templateWorkbook = new Workbook(templatePath);

                // Enable multi‑thread reading for the cells collection (optional but recommended)
                if (templateWorkbook.Worksheets.Count > 0)
                    templateWorkbook.Worksheets[0].Cells.MultiThreadReading = true;

                // Create a WorkbookDesigner instance for this thread
                WorkbookDesigner designer = new WorkbookDesigner(templateWorkbook);

                // Set the data source for smart markers
                designer.SetDataSource("Person", dataSource);

                // Process the smart markers
                designer.Process();

                // Store the processed workbook so the main thread can merge it later
                ProcessedWorkbooks[templateIndex] = designer.Workbook;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing template index {templateIndex}: {ex.Message}");
            }
        }
    }
}