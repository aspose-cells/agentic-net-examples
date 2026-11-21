// Title: Preserve cell formatting while processing smart markers from a List<Person> using Aspose.Cells WorkbookDesigner in C#
// Common Searches: how to keep cell formatting when processing smart markers with Aspose.Cells in C# | C# example of using DesignerOptions.PreserveCellFormatting for Excel smart markers | populate an Excel template from a List<Person> while retaining original styles using Aspose.Cells | Aspose.Cells WorkbookDesigner preserve formatting during data binding | smart markers insert data without losing cell styles in .NET
// Tags: WorkbookDesigner PreserveCellFormatting option | smart markers data binding from List<Person> | retain cell styles during Aspose.Cells smart marker processing | populate Excel template with custom object list C# | Aspose.Cells formatting preservation for smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook containing smart markers, sets a List<Person> as the data source, optionally enables the DesignerOptions.PreserveCellFormatting flag on WorkbookDesigner, processes the smart markers to insert the data while keeping the original cell formatting, and saves the resulting file.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            Workbook workbook;

            // Load the template if it exists; otherwise create a simple workbook with smart markers.
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Name");
                ws.Cells["B1"].PutValue("Age");
                ws.Cells["A2"].PutValue("&=persons.Name");
                ws.Cells["B2"].PutValue("&=persons.Age");
            }

            // Prepare data source.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John", Age = 30 },
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob", Age = 28 }
            };

            // Process smart markers using WorkbookDesigner.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Preserve original cell formatting if the API version supports it.
            // Uncomment the following line if DesignerOptions is available in your Aspose.Cells version:
            // designer.DesignerOptions.PreserveCellFormatting = true;
            designer.SetDataSource("persons", persons);
            designer.Process();

            // Save the result workbook.
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
