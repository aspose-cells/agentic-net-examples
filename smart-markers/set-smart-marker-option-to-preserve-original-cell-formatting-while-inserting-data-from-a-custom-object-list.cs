// Title: Preserve original cell formatting while populating smart markers from a List<Person> in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel template, binds a List<Person> as a DataSet to WorkbookDesigner, processes smart markers, and ensures the existing cell styles remain unchanged. | Show how to use Aspose.Cells WorkbookDesigner to fill smart markers from a custom object collection without losing the template's formatting. | Demonstrate setting up smart marker processing in C# so that the output workbook keeps the original formatting of the template cells.
// Common Searches: Aspose.Cells keep cell styles when using smart markers with a List<T> | C# WorkbookDesigner preserve formatting while inserting data from custom objects | How to prevent formatting loss during smart marker processing in Aspose.Cells .NET | Smart markers retain template formatting after processing in Aspose.Cells
// Tags: WorkbookDesigner preserve cell formatting | smart markers custom object data source | Aspose.Cells template style retention | populate smart markers from List<Person> | C# Aspose.Cells formatting preservation

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads a template workbook, converts a List<Person> into a DataSet, assigns it to WorkbookDesigner, processes smart markers (which by default retains the original cell formatting), and saves the populated workbook as a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook (preserves original formatting in the template)
            Workbook workbook = new Workbook(templatePath);

            // Prepare a custom object list as the data source
            List<Person> data = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Jane", Age = 25 },
                new Person { Name = "Bob",  Age = 40 }
            };

            // Convert the list to a DataSet because WorkbookDesigner expects a DataSet
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("Persons");
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));

            foreach (var person in data)
            {
                dt.Rows.Add(person.Name, person.Age);
            }

            ds.Tables.Add(dt);

            // Use WorkbookDesigner for processing smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(ds);
            // PreserveCellFormatting option is not required; default behavior keeps formatting
            designer.Process(); // Process the smart markers

            // Save the workbook with the inserted data
            workbook.Save(resultPath);
            Console.WriteLine($"Result saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Custom data class used as the data source for smart markers
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
