// Title: Process multiple Excel templates with smart markers in parallel using separate WorkbookDesigner instances and merge results with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads several Excel template files, creates an individual WorkbookDesigner for each, binds distinct data sources, executes Process() concurrently with Parallel.For, and merges all processed worksheets into one workbook. | Show how to safely combine workbooks when running parallel smart‑marker processing, including lock usage and enabling MultiThreadReading on worksheets. | Provide an example that checks for missing template files and handles exceptions while performing concurrent WorkbookDesigner processing and saving the final merged workbook.
// Common Searches: how to use Aspose.Cells WorkbookDesigner for parallel smart marker processing in C# | merge worksheets after concurrent processing with Aspose.Cells .NET | concurrent way to combine multiple Excel workbooks using Aspose.Cells | enable MultiThreadReading for cells when processing smart markers in parallel | process several Excel templates with different data sources simultaneously Aspose.Cells
// Tags: parallel workbookdesigner processing Aspose.Cells | smart marker merging multiple templates | concurrent workbook combine Aspose.Cells | multithread reading cells Aspose.Cells | c# aspose.cells concurrent template processing

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads multiple Excel template files containing smart markers, creates a separate WorkbookDesigner for each template, binds a unique List<Person> data source, processes the smart markers concurrently with Parallel.For (enabling MultiThreadReading), and merges the resulting worksheets into a single Workbook using a lock for safe concurrency before saving as MergedResult.xlsx. It also includes checks for missing files and exception handling.
class Program
{
    static void Main()
    {
        // Paths to the template workbooks (each contains smart markers)
        string[] templates = { "Template1.xlsx", "Template2.xlsx", "Template3.xlsx" };

        // Example data sources – one per template
        var dataSources = new List<object>
        {
            new List<Person>
            {
                new Person("John", 30),
                new Person("Alice", 25)
            },
            new List<Person>
            {
                new Person("Bob", 40),
                new Person("Eve", 35)
            },
            new List<Person>
            {
                new Person("Mike", 28),
                new Person("Sara", 32)
            }
        };

        // Workbook that will hold the merged result
        Workbook finalWorkbook = new Workbook();
        finalWorkbook.Worksheets.Clear(); // start with no sheets

        object mergeLock = new object();

        // Process each template in parallel
        Parallel.For(0, templates.Length, i =>
        {
            try
            {
                string path = templates[i];

                // Verify that the template file exists before loading
                if (!File.Exists(path))
                {
                    Console.WriteLine($"Warning: Template file not found – skipping: {path}");
                    return;
                }

                // Load the template workbook
                Workbook templateWb = new Workbook(path);

                // Enable multi‑thread reading for the cells collection (optional but safe)
                if (templateWb.Worksheets.Count > 0)
                {
                    templateWb.Worksheets[0].Cells.MultiThreadReading = true;
                }

                // Create a WorkbookDesigner for this workbook
                WorkbookDesigner designer = new WorkbookDesigner(templateWb);

                // Bind the corresponding data source (name can be any identifier used in the template)
                designer.SetDataSource("Data", dataSources[i]);

                // Process smart markers
                designer.Process();

                // Merge the processed workbook into the final workbook
                lock (mergeLock)
                {
                    if (finalWorkbook.Worksheets.Count == 0)
                    {
                        // First processed workbook becomes the base
                        finalWorkbook.Copy(templateWb);
                    }
                    else
                    {
                        // Add all worksheets from the processed workbook
                        finalWorkbook.Combine(templateWb);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing template '{templates[i]}': {ex.Message}");
            }
        });

        // Save the merged workbook if at least one worksheet exists
        if (finalWorkbook.Worksheets.Count > 0)
        {
            try
            {
                finalWorkbook.Save("MergedResult.xlsx");
                Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("No templates were processed. Merged workbook was not created.");
        }
    }

    // Simple POCO class used as a data source for smart markers
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }
}
