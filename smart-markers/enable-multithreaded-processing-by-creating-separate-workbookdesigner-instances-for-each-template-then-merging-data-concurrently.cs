// Title: Concurrent Excel template processing with Aspose.Cells WorkbookDesigner (C#)
// Description: Loads multiple Excel templates, enables MultiThreadReading, creates a separate WorkbookDesigner for each, binds a list of Person objects as the "Data" source, processes smart markers in parallel tasks, and merges the successful workbooks into a single file (MergedResult.xlsx).
// Keywords: Aspose.Cells multithreaded processing | WorkbookDesigner parallel execution | smart markers C# | merge multiple workbooks Aspose | MultiThreadReading cells | concurrent Excel template generation | Aspose.Cells combine workbooks
// Common Searches: process Excel templates concurrently with Aspose.Cells | parallel smart marker processing C# | merge workbooks after parallel execution Aspose | enable MultiThreadReading for WorkbookDesigner | combine multiple WorkbookDesigner results
// Developer Intent: Run each Excel template on its own thread using WorkbookDesigner, then combine the processed workbooks into one file.
// Use Cases: Generate department‑level reports simultaneously and produce a master workbook. | Create a batch of invoices from different templates in parallel, then archive them together. | Aggregate regional sales data from several smart‑marker templates concurrently for executive review.
// AI Prompts: Write C# code that uses Aspose.Cells to process a collection of Excel templates with smart markers on separate threads and merges the outputs into a single workbook. | Explain the performance impact of MultiThreadReading when using WorkbookDesigner and list thread‑safety best practices. | Show how to log errors for individual template tasks while still consolidating successfully processed workbooks.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

// Loads multiple Excel templates, enables MultiThreadReading, creates a separate WorkbookDesigner for each, binds a list of Person objects as the "Data" source, processes smart markers in parallel tasks, and merges the successful workbooks into a single file (MergedResult.xlsx).
class MultiThreadWorkbookDesignerDemo
{
    static void Main()
    {
        // Define template files and their corresponding data sources.
        var templates = new List<(string templatePath, object dataSource)>
        {
            ("Template1.xlsx", GetSampleData1()),
            ("Template2.xlsx", GetSampleData2()),
            ("Template3.xlsx", GetSampleData3())
        };

        var processingTasks = new List<Task<Workbook>>();

        // Process each template in its own task.
        foreach (var item in templates)
        {
            processingTasks.Add(Task.Run(() =>
            {
                try
                {
                    // Verify that the template file exists.
                    if (!File.Exists(item.templatePath))
                        throw new FileNotFoundException($"Template file not found: {item.templatePath}");

                    // Load the template workbook.
                    Workbook wb = new Workbook(item.templatePath);

                    // Enable multi‑thread reading for the worksheet's cells.
                    wb.Worksheets[0].Cells.MultiThreadReading = true;

                    // Create a WorkbookDesigner for this workbook.
                    WorkbookDesigner designer = new WorkbookDesigner(wb);

                    // Bind the data source (the name "Data" matches the smart markers in the template).
                    designer.SetDataSource("Data", item.dataSource);

                    // Process smart markers.
                    designer.Process();

                    // Return the processed workbook.
                    return designer.Workbook;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{item.templatePath}': {ex.Message}");
                    return null;
                }
            }));
        }

        try
        {
            // Wait for all parallel tasks to complete.
            Task.WaitAll(processingTasks.ToArray());
        }
        catch (AggregateException aggEx)
        {
            foreach (var ex in aggEx.InnerExceptions)
                Console.WriteLine($"Task error: {ex.Message}");
        }

        // Merge all successfully processed workbooks into a single workbook.
        Workbook finalWorkbook = null;
        foreach (var task in processingTasks)
        {
            Workbook processed = task.Result;
            if (processed == null)
                continue; // Skip failed tasks.

            if (finalWorkbook == null)
            {
                // Use the first processed workbook as the base.
                finalWorkbook = processed;
            }
            else
            {
                // Combine subsequent workbooks into the base workbook.
                finalWorkbook.Combine(processed);
            }
        }

        // Save the merged result if at least one workbook was processed.
        if (finalWorkbook != null)
        {
            try
            {
                finalWorkbook.Save("MergedResult.xlsx");
                Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving merged workbook: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("No workbooks were processed; merged file not created.");
        }
    }

    // Sample data for the first template.
    static List<Person> GetSampleData1()
    {
        return new List<Person>
        {
            new Person("Alice", 28),
            new Person("Bob", 35)
        };
    }

    // Sample data for the second template.
    static List<Person> GetSampleData2()
    {
        return new List<Person>
        {
            new Person("Charlie", 22),
            new Person("Diana", 31)
        };
    }

    // Sample data for the third template.
    static List<Person> GetSampleData3()
    {
        return new List<Person>
        {
            new Person("Eve", 27),
            new Person("Frank", 40)
        };
    }

    // Simple POCO class used as a data source for smart markers.
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
