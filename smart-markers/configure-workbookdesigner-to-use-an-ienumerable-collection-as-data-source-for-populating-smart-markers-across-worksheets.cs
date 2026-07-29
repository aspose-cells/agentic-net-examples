// Title: Use IEnumerable with WorkbookDesigner to Fill Smart Markers on Multiple Sheets (Aspose.Cells .NET)
// Description: Demonstrates how to bind a List<Person> (IEnumerable) to WorkbookDesigner, place smart markers (e.g., &=Person.Name, &=Person.Age, &=Person.Count) on two worksheets, process all markers, and save the populated Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | WorkbookDesigner | IEnumerable | List<Person> | smart markers | C# | .NET | multiple worksheets | SetDataSource | Count smart marker | Excel automation | populate workbook
// Common Searches: Aspose.Cells bind List<T> to WorkbookDesigner | smart markers from IEnumerable collection | WorkbookDesigner SetDataSource example C# | populate Excel with List<Person> using Aspose | smart marker Count with collection Aspose.Cells | multiple sheet smart markers Aspose.Cells
// Developer Intent: Learn how to bind an IEnumerable collection to WorkbookDesigner and generate smart‑marker populated worksheets across several sheets in Aspose.Cells for .NET.
// Use Cases: Create an employee roster sheet that lists each person's Name and Age from a List<Person> using smart markers. | Add a summary sheet that shows the total number of employees with the &=Person.Count marker. | Reuse a single IEnumerable data source to populate smart markers on multiple worksheets within one workbook.
// AI Prompts: Show me how to bind a DataTable to WorkbookDesigner for smart‑marker processing in Aspose.Cells. | Explain how to handle nested collections (e.g., List<Order> inside Person) with WorkbookDesigner smart markers. | Give best practices for troubleshooting missing or mismatched smart markers when using IEnumerable with WorkbookDesigner.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple POCO class that will be used as the data source
    // Demonstrates how to bind a List<Person> (IEnumerable) to WorkbookDesigner, place smart markers (e.g., &=Person.Name, &=Person.Age, &=Person.Count) on two worksheets, process all markers, and save the populated Excel file using Aspose.Cells for .NET.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Prepare an IEnumerable collection (List<Person>) as the data source
                List<Person> people = new List<Person>
                {
                    new Person("Alice", 28),
                    new Person("Bob", 35),
                    new Person("Charlie", 42)
                };

                // Create a new workbook and add two worksheets
                Workbook workbook = new Workbook();

                // First worksheet (default worksheet at index 0)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Employees";

                // Add smart markers to the first worksheet
                sheet1.Cells["A1"].PutValue("Name");
                sheet1.Cells["B1"].PutValue("Age");
                sheet1.Cells["A2"].PutValue("&=Person.Name");
                sheet1.Cells["B2"].PutValue("&=Person.Age");

                // Add second worksheet and obtain its reference
                int sheet2Index = workbook.Worksheets.Add();               // Returns the index of the new sheet
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Summary";

                // Add smart markers to the second worksheet
                sheet2.Cells["A1"].PutValue("Total Employees");
                sheet2.Cells["B1"].PutValue("&=Person.Count"); // Count works on IEnumerable

                // Initialize the WorkbookDesigner with the workbook instance
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Bind the IEnumerable collection to the variable name "Person"
                designer.SetDataSource("Person", people);

                // Process all smart markers in the workbook
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkersFromIEnumerable.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the populated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
