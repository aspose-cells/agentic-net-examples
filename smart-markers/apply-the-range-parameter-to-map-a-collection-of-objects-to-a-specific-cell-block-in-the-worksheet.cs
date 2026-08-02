// Title: Map a List of objects to a specific Excel range using Aspose.Cells (C#)
// Description: Demonstrates how to create a Workbook, define a Range starting at A2 sized to a List<Person>, fill a 2‑D array with Name and Age values, assign it to Range.Value, add headers, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# range mapping | populate Excel from List | CreateRange method Aspose | assign 2D array to cells | Excel smart markers alternative | write collection to worksheet | C# Excel export tutorial | Range.Value property | dynamic Excel table C# | Aspose.Cells example
// Common Searches: Aspose.Cells map List to range | C# create Excel range from collection | How to use CreateRange in Aspose.Cells | Set Range.Value with 2D array Aspose | Add headers to Excel range Aspose.Cells | Export List<Person> to Excel using Aspose
// Developer Intent: I need to write a collection of objects into a predefined block of cells in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Generate an employee roster where each row corresponds to a Person object and the table starts at cell A2. | Export sales or inventory data to a dynamically sized table that begins at a specific address. | Create a reusable helper that writes any IList<T> to a worksheet range with optional column headers. | Build a dynamic report that can be merged with other sheets after populating data.
// AI Prompts: Write a generic C# method that takes IList<T>, a start cell address, and a Worksheet, then uses reflection to fill a Range with the objects' property values. | Show how to achieve the same data mapping using Aspose.Cells Smart Markers instead of a manual 2‑D array. | Provide robust error handling for collections that exceed worksheet limits or contain null values when assigning to a Range. | Explain how to apply styling (borders, header formatting) to the generated range after populating data with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsRangeMappingDemo
{
    // Sample data class
    // Demonstrates how to create a Workbook, define a Range starting at A2 sized to a List<Person>, fill a 2‑D array with Name and Age values, assign it to Range.Value, add headers, and save the file with Aspose.Cells for .NET.
    public class Person
    {
        public string? Name { get; set; }   // Made nullable to satisfy non‑nullable warning
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample collection of objects to map
                List<Person> people = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                };

                // Determine the size of the range (rows = number of items, columns = 2 for Name and Age)
                int totalRows = people.Count;
                int totalColumns = 2; // Name and Age

                // Create a range starting at cell A2 (row index 1, column index 0)
                Aspose.Cells.Range dataRange = cells.CreateRange(1, 0, totalRows, totalColumns);

                // Prepare a 2‑dimensional array matching the range dimensions
                object[,] values = new object[totalRows, totalColumns];
                for (int i = 0; i < totalRows; i++)
                {
                    values[i, 0] = people[i].Name; // First column: Name
                    values[i, 1] = people[i].Age;  // Second column: Age
                }

                // Assign the array to the range; Aspose.Cells will populate the cells accordingly
                dataRange.Value = values;

                // Optional: add headers above the data
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Age");

                // Save the workbook to a file
                string outputPath = "PeopleData.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
