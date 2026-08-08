// Title: Implement ICustomTypeProvider for C# Smart Markers in Aspose.Cells
// Description: A step‑by‑step C# example that creates an Excel workbook, places smart markers, and binds a custom POCO (Person) via ICustomTypeProvider to WorkbookDesigner for dynamic data population. Works with Aspose.Cells 23+ worldwide.
// Keywords: Aspose.Cells | C# smart markers | ICustomTypeProvider | WorkbookDesigner | Excel data binding | custom POCO | Excel generation | smart marker example | GitHub sample | global
// Common Searches: how to bind a custom C# class to Aspose.Cells smart markers | ICustomTypeProvider example for Aspose.Cells | Aspose.Cells smart markers tutorial C# | populate Excel with POCO using WorkbookDesigner | Aspose.Cells custom type provider GitHub
// Developer Intent: Create an Excel file where smart markers are automatically replaced with values from a user‑defined C# class using Aspose.Cells.
// Use Cases: Insert smart markers like &=Person.FirstName, &=Person.LastName, &=Person.Age into a worksheet and have them resolved at runtime. | Expose additional calculated properties through ICustomTypeProvider without modifying the original POCO. | Bind the custom type to WorkbookDesigner with SetDataSource and generate the final workbook in a single call.
// AI Prompts: Generate C# code that implements ICustomTypeProvider for a Person class and uses the new properties in Aspose.Cells smart markers. | Explain how to register a custom type provider with WorkbookDesigner to enable advanced smart marker binding. | Show a complete GitHub‑ready example that creates an Excel file, adds smart markers, and processes them using ICustomTypeProvider.

using System;
using Aspose.Cells; // Core Aspose.Cells classes

// Simple POCO class used as data source for smart markers.
// A step‑by‑step C# example that creates an Excel workbook, places smart markers, and binds a custom POCO (Person) via ICustomTypeProvider to WorkbookDesigner for dynamic data population. Works with Aspose.Cells 23+ worldwide.
public class Person
{
    public string FirstName { get; set; }
    public string LastName  { get; set; }
    public int    Age       { get; set; }
}

public class SmartMarkerWithCustomTypeProviderDemo
{
    public static void Run()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and place smart markers that reference
            //    the custom object's properties.
            // ------------------------------------------------------------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Smart markers use the syntax "&=ObjectName.PropertyName"
            sheet.Cells["A1"].PutValue("&=Person.FirstName");
            sheet.Cells["A2"].PutValue("&=Person.LastName");
            sheet.Cells["A3"].PutValue("&=Person.Age");

            // ------------------------------------------------------------
            // 2. Create an instance of the custom class and populate data.
            // ------------------------------------------------------------
            var person = new Person
            {
                FirstName = "John",
                LastName  = "Doe",
                Age       = 30
            };

            // ------------------------------------------------------------
            // 3. Use WorkbookDesigner to process the smart markers.
            // ------------------------------------------------------------
            var designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // The name "Person" must match the prefix used in the smart markers.
            designer.SetDataSource("Person", person);

            // Process the markers (false = do not preserve empty rows/columns).
            designer.Process(false);

            // ------------------------------------------------------------
            // 4. Save the resulting workbook.
            // ------------------------------------------------------------
            const string outputPath = "SmartMarkerWithCustomTypeProvider.xlsx";
            designer.Workbook.Save(outputPath);
            Console.WriteLine($"Workbook created: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing: {ex.Message}");
        }
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        SmartMarkerWithCustomTypeProviderDemo.Run();
    }
}
