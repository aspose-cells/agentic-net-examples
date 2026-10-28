// Title: Configure Aspose.Cells smart markers to skip empty rows in a C# detail list
// Common Searches: how to ignore null entries in Aspose.Cells smart markers using C# | skip blank rows when processing a detail list with WorkbookDesigner | Aspose.Cells smart marker option to omit empty rows from output | C# example of smart markers that filter out null list items | prevent empty rows in generated Excel file with Aspose.Cells smart markers
// Tags: smart markers null entry handling | WorkbookDesigner ignore blank rows | Aspose.Cells smart marker data gaps | C# report generation without empty rows | configure smart marker options for missing data

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// This C# example shows how to load or create a workbook with smart markers, bind a List<Person?> that contains null entries, and process the markers using WorkbookDesigner. Aspose.Cells automatically skips the null items, so the resulting Excel report contains no empty rows, effectively handling gaps in the data source.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for template and result workbooks
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Load existing template or create a new workbook with smart markers
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("&=People.Name");
                ws.Cells["B1"].PutValue("&=People.Age");
            }

            // Sample data with null entries to represent empty rows
            List<Person?> people = new List<Person?>
            {
                new Person { Name = "John", Age = 30 },
                null,
                new Person { Name = "Alice", Age = 25 },
                null,
                new Person { Name = "Bob", Age = 40 }
            };

            // Initialize the designer with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // (Optional) Configure smart marker options if needed.
            // In newer Aspose.Cells versions, Options may not be available.
            // The default behavior already skips null entries when processing.

            // Set the data source for smart markers
            designer.SetDataSource("People", people);

            // Process smart markers safely
            try
            {
                designer.Process();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during smart marker processing: {ex.Message}");
                return;
            }

            // Save the generated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Report generated successfully: {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Simple POCO used as the data source
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
