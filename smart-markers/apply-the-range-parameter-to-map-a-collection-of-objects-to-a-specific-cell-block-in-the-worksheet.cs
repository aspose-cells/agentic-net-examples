// Title: Map a List<Person> to a defined Excel range starting at A2 with headers using Aspose.Cells for .NET
// AI Prompts: Write C# code that converts a List<Person> into a 2‑dimensional object array and assigns it to a worksheet range beginning at cell A2 with a header row using Aspose.Cells. | Show how to compute the lower‑right cell address for a dynamic range based on the collection count with CellsHelper and create the corresponding Aspose.Cells Range. | Demonstrate setting the Value property of an Aspose.Cells Range to a 2‑D array and saving the workbook to an .xlsx file.
// Common Searches: asp.net map list of objects to excel range using aspose.cells | c# create dynamic range starting at A2 based on collection size aspose.cells | populate excel worksheet with object array asp.net aspose.cells | how to assign a 2d array to a range in aspose.cells c#
// Tags: assign object list to Excel range Aspose.Cells | dynamic range calculation CellsHelper Aspose.Cells | populate worksheet with 2d array .NET | map collection to Excel cells Aspose.Cells | create range with header row Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsRangeMappingDemo
{
    // Simple data model
    // The example creates a workbook, builds a 2‑D object array from a List<Person> (including a header row), calculates a dynamic target range that starts at A2, assigns the array to that Aspose.Cells Range, and saves the file as PeopleRangeMapping.xlsx.
    public class Person
    {
        public string Name { get; set; } = null!;
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

                // Determine the size of the target range (including header row)
                int totalRows = people.Count + 1; // +1 for header
                int totalColumns = 2; // Name and Age

                // Create a 2‑dimensional array that matches the range dimensions
                object[,] data = new object[totalRows, totalColumns];

                // Fill header
                data[0, 0] = "Name";
                data[0, 1] = "Age";

                // Fill data rows from the collection
                for (int i = 0; i < people.Count; i++)
                {
                    data[i + 1, 0] = people[i].Name;
                    data[i + 1, 1] = people[i].Age;
                }

                // Define the target range (starting at cell A2)
                // Upper‑left cell: A2, lower‑right cell calculated from size
                string upperLeft = "A2";
                // CellsHelper uses zero‑based indices, so add 1 to row index for the header offset
                string lowerRight = CellsHelper.CellIndexToName(totalRows, totalColumns - 1);
                Aspose.Cells.Range targetRange = cells.CreateRange(upperLeft, lowerRight);

                // Assign the 2‑D array to the range
                targetRange.Value = data;

                // Save the workbook
                string outputPath = "PeopleRangeMapping.xlsx";
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
