// Title: Create a custom ICustomTypeProvider in C# to expose computed Person properties for Aspose.Cells smart markers
// Common Searches: how to bind computed properties like FullName to Aspose.Cells smart markers using ICustomTypeProvider | C# example of custom type provider for smart markers with Aspose.Cells WorkbookDesigner | expose additional fields from a .NET class to smart markers in an Excel file using Aspose.Cells | use ICustomTypeProvider to make calculated properties available in Aspose.Cells smart marker templates | Aspose.Cells smart marker custom type provider tutorial for .NET developers
// Tags: custom ICustomTypeProvider for Aspose.Cells smart markers | binding computed properties to smart markers C# | WorkbookDesigner data source with custom type provider | expose FullName and AgeInMonths in Excel using Aspose.Cells | smart marker template with additional class fields

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample defines a Person class with computed FullName and AgeInMonths properties, inserts smart marker cells referencing those properties, sets a List<Person> as the data source via WorkbookDesigner, processes the markers using a custom ICustomTypeProvider, and saves the populated workbook as output.xlsx.
public class Person
{
    // Nullable strings to satisfy non‑nullable warnings.
    public string? FirstName { get; set; }
    public string? LastName  { get; set; }
    public int Age          { get; set; }

    // Computed property: full name.
    public string FullName => $"{FirstName} {LastName}";

    // Computed property: age in months.
    public int AgeInMonths => Age * 12;
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Insert sample smart marker cells (optional, for demonstration).
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("&=Persons.FullName");
            sheet.Cells["A2"].PutValue("&=Persons.AgeInMonths");

            // Prepare a data source.
            var persons = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 }
            };

            // Use WorkbookDesigner to process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Persons", persons);
            designer.Process(); // Replaces smart markers with data.

            // Save the resulting workbook.
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
