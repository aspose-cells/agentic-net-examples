// Title: Expose a Computed Property for Aspose.Cells Smart Markers – Person.ExtraInfo Example
// Description: Demonstrates how to add a virtual property (ExtraInfo) to a C# data class and bind a List<Person> to smart markers using WorkbookDesigner, allowing the markers &=$People.Name and &=$People.ExtraInfo to render combined name‑age values in an Excel workbook.
// Keywords: Aspose.Cells smart markers | C# computed property | ExtraInfo virtual field | WorkbookDesigner data binding | ICustomTypeProvider alternative | .NET Excel export | dynamic smart marker property | Excel automation C# | Aspose.Cells example GitHub
// Common Searches: how to show a calculated field in Aspose.Cells smart markers | bind a list of objects with extra properties to smart markers in C# | use computed property with Aspose.Cells WorkbookDesigner | smart marker reference virtual property Aspose.Cells | C# example for adding ExtraInfo to smart marker data source
// Developer Intent: Add virtual or computed members to data objects so they can be accessed directly from Aspose.Cells smart markers.
// Use Cases: Display combined name and age information in a single cell without modifying the original data model. | Provide read‑only, on‑the‑fly values for reporting templates that use smart markers. | Simplify Excel‑based reports by exposing derived fields through a custom type provider or computed property.
// AI Prompts: Generate a C# class that implements ICustomTypeProvider to expose a virtual property ExtraInfo for Person objects used in Aspose.Cells smart markers. | Show how to register a custom type provider with WorkbookDesigner and bind a List<Person> to the "People" smart marker collection. | Provide sample code that processes smart markers referencing both real and computed properties and saves the resulting Excel file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCustomType
{
    // Data class used as the data source for smart markers.
    // Includes a computed property ExtraInfo that will be accessed by the smart markers.
    // Demonstrates how to add a virtual property (ExtraInfo) to a C# data class and bind a List<Person> to smart markers using WorkbookDesigner, allowing the markers &=$People.Name and &=$People.ExtraInfo to render combined name‑age values in an Excel workbook.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        // Computed property exposed to smart markers as $People.ExtraInfo.
        public string ExtraInfo => $"Name: {Name}, Age: {Age}";
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define smart markers that reference the regular and the computed property.
                sheet.Cells["A1"].PutValue("&=$People.Name");
                sheet.Cells["A2"].PutValue("&=$People.ExtraInfo");

                // Prepare the data source – a list of Person objects.
                List<Person> people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 25 }
                };

                // Use WorkbookDesigner to process smart markers.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };

                // Bind the list to the name "People" used in the smart markers.
                designer.SetDataSource("People", people);

                // Process the smart markers (false = do not preserve smart markers after processing).
                designer.Process(false);

                // Save the result.
                string outputPath = "SmartMarkerCustomTypeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
