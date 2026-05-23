using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        const string templatePath = "TemplateWithHiddenSheets.xlsx";
        const string resultPath = "Result.xlsx";

        try
        {
            // Verify that the template file exists before loading
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {Path.GetFullPath(templatePath)}");
                return;
            }

            // Load the workbook that contains hidden worksheets with smart markers
            Workbook workbook = new Workbook(templatePath);

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a sample data source (replace with your actual data)
            var people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 }
            };

            // Bind the data source to a smart marker name
            designer.SetDataSource("People", people);

            // Process all smart markers in the workbook (including hidden sheets)
            designer.Process();

            // Unhide any worksheets that were hidden before processing
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!sheet.IsVisible)
                {
                    sheet.IsVisible = true;
                }
            }

            // Save the processed workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(resultPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Sample data class used for the data source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}