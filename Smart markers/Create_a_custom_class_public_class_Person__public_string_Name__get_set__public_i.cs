using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom data class
public class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}

public class LoadAndImportExample
{
    public static void Main()
    {
        // Load an existing XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile); // uses Workbook(string) constructor
        Worksheet sheet = workbook.Worksheets[0];

        // Sample list of Person objects to import
        List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 }
        };

        // Specify which properties to import
        string[] propertyNames = { "Name", "Age" };

        // Import the custom objects into the worksheet starting at cell A1 (row 0, column 0)
        // Show property names in the first row, insert rows if needed, and convert strings to numbers where applicable
        sheet.Cells.ImportCustomObjects(
            people,                // ICollection list
            propertyNames,         // string[] propertyNames
            true,                  // isPropertyNameShown
            0,                     // firstRow
            0,                     // firstColumn
            people.Count,          // rowNumber
            true,                  // insertRows
            "",                    // dateFormatString (not used for Person)
            true                   // convertStringToNumber
        );

        // Save the modified workbook
        workbook.Save("output.xlsx"); // uses Workbook.Save(string) method
    }
}