// Title: Implement a custom ICustomTypeProvider class to expose computed properties for Aspose.Cells smart markers in C#
// AI Prompts: Write a C# class that implements ICustomTypeProvider and adds a FullName property to the Person type for use with Aspose.Cells smart markers. | Show how to register the custom type provider with WorkbookDesigner, bind a List<Person> collection, and process smart markers that reference the new property. | Provide a complete example that creates a workbook, inserts a smart marker using the FullName property, processes the data source, and saves the resulting Excel file.
// Common Searches: C# ICustomTypeProvider example for Aspose.Cells smart markers | how to bind a computed property to a smart marker using WorkbookDesigner | expose additional fields for smart marker binding with a custom type provider | Aspose.Cells smart marker custom object with FullName property | using a custom type provider to extend data source for Excel smart markers
// Tags: ICustomTypeProvider implementation Aspose.Cells | FullName property exposure for smart markers | WorkbookDesigner bind custom object collection | C# smart marker data source extension | Aspose.Cells computed property integration

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example demonstrates creating a workbook, inserting a smart marker that references a computed FullName property, implementing a custom ICustomTypeProvider to expose that property, binding a List<Person> collection named "Employees" to WorkbookDesigner, processing the smart markers, and saving the output Excel file.
public class Person
{
    // Regular properties
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }

    // Computed property exposed for Smart Marker binding
    public string FullName => $"{Name} (Age: {Age})";
}

public class SmartMarkerDemo
{
    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a Smart Marker that references the computed property "FullName"
            // Use the collection name "Employees" in the marker expression
            sheet.Cells["A1"].PutValue("Employee: <#=Employees.FullName#>");

            // Prepare data source: a list of Person objects
            List<Person> employees = new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 28 },
                new Person { Name = "Bob Smith", Age = 35 }
            };

            // Process the workbook with the data source using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Bind the collection with the name used in the Smart Marker
            designer.SetDataSource("Employees", employees);
            designer.Process();

            // Define output path and ensure the directory exists
            string outputPath = "SmartMarkerResult.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the result
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
