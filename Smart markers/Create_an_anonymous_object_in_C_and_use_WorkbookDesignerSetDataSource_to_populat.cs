using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Determine the full path to the template file (assumed to be in the same folder as the executable)
        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SmartMarkerTemplate.xlsx");

        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Load the Excel template that contains smart markers (e.g., &Person.Name, &Person.Age)
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook(templatePath);

        // Create an anonymous object whose property names match the smart markers
        var person = new { Name = "John Doe", Age = 35 };

        // Bind the anonymous object to the variable name used in the template
        designer.SetDataSource("Person", person);

        // Process the smart markers and populate the worksheet with data from the anonymous object
        designer.Process();

        // Save the resulting workbook
        string resultPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Result.xlsx");
        designer.Workbook.Save(resultPath);

        Console.WriteLine($"Result saved to: {resultPath}");
    }
}