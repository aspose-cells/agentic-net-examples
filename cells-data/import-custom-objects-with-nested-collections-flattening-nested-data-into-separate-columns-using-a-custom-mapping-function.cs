// Title: Flatten a list of Person objects with multiple Address entries into separate Excel columns using Aspose.Cells ImportObjectArray in C#
// AI Prompts: Generate C# code that iterates over a List<Person>, extracts each person's name, age, and up to two address fields, builds a two‑dimensional object array, and writes it to an XLSX file with Cells.ImportObjectArray. | Create a mapping method for Aspose.Cells that converts a Person instance containing a List<Address> into a flat object[] suitable for ImportObjectArray, handling missing addresses gracefully. | Show how to construct a header row, use Cells.ImportObjectArray to export the flattened Person data to a worksheet, and then save the workbook as an Excel file.
// Common Searches: how to export a list of objects with nested collections to Excel using Aspose.Cells C# | Aspose.Cells ImportObjectArray flatten nested list of addresses example | C# map Person class with multiple Address objects to separate Excel columns | write custom mapping function for ImportObjectArray with Aspose.Cells
// Tags: importobjectarray flatten nested collections c# | aspocells export nested objects to xlsx | custom mapping function for excel export c# | person address flattening aspocells | c# write header row importobjectarray

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The example creates a workbook, generates sample Person objects each with a list of Address objects, builds a header row and flat rows by extracting up to two address fields per person via a custom MapPersonToRow method, imports the resulting object array into the worksheet using Cells.ImportObjectArray, and saves the file as FlattenedData.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data: list of Person objects with nested Addresses
        List<Person> people = GetSampleData();

        // Flatten data using a custom mapping function
        List<object[]> rows = new List<object[]>();

        // Header row
        rows.Add(new object[] { "Name", "Age", "Address1", "City1", "Address2", "City2" });

        // Convert each Person to a flat row
        foreach (var person in people)
        {
            rows.Add(MapPersonToRow(person));
        }

        // Import the flattened rows into the worksheet starting at cell A1
        cells.ImportObjectArray(rows.ToArray(), 0, 0, true);

        // Save the workbook
        workbook.Save("FlattenedData.xlsx");
    }

    // Mapping function: converts a Person object into a flat object array
    static object[] MapPersonToRow(Person p)
    {
        // Prepare placeholders for up to two addresses (adjust as needed)
        string address1 = "";
        string city1 = "";
        string address2 = "";
        string city2 = "";

        if (p.Addresses != null && p.Addresses.Count > 0)
        {
            address1 = p.Addresses[0].Street;
            city1 = p.Addresses[0].City;
        }
        if (p.Addresses != null && p.Addresses.Count > 1)
        {
            address2 = p.Addresses[1].Street;
            city2 = p.Addresses[1].City;
        }

        return new object[] { p.Name, p.Age, address1, city1, address2, city2 };
    }

    // Sample data generator
    static List<Person> GetSampleData()
    {
        return new List<Person>
        {
            new Person
            {
                Name = "John Doe",
                Age = 30,
                Addresses = new List<Address>
                {
                    new Address { Street = "123 Main St", City = "New York" },
                    new Address { Street = "456 Oak Ave", City = "Boston" }
                }
            },
            new Person
            {
                Name = "Jane Smith",
                Age = 25,
                Addresses = new List<Address>
                {
                    new Address { Street = "789 Pine Rd", City = "Chicago" }
                }
            }
        };
    }
}

// Custom data classes
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
    public List<Address> Addresses { get; set; }
}

public class Address
{
    public string Street { get; set; }
    public string City { get; set; }
}
