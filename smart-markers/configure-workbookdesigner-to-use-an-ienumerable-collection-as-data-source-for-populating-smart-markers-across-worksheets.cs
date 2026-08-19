// Title: Bind an IEnumerable<T> to WorkbookDesigner for Smart Markers on Multiple Worksheets (C#)
// Description: Shows how to use a List<Person> (IEnumerable) as the data source for Aspose.Cells WorkbookDesigner, place smart markers on two sheets, process them, and save the workbook.
// Keywords: Aspose.Cells | WorkbookDesigner | IEnumerable | C# smart markers | List<T> data source | multiple worksheets | variable smart marker | populate Excel from collection | Aspose.Cells example
// Common Searches: Aspose.Cells bind List to WorkbookDesigner | smart markers using IEnumerable C# | multiple sheet smart markers Aspose.Cells | variable smart marker first item Aspose | populate Excel with collection Aspose.Cells
// Developer Intent: Bind an IEnumerable collection to WorkbookDesigner and generate smart‑marker populated worksheets.
// Use Cases: Create an employee roster sheet by iterating over a List<Person> with smart markers. | Display a specific property (e.g., first employee name) on a summary sheet using a variable smart marker. | Generate a multi‑sheet report workbook from a single data source without manual loops. | Reuse the same collection across several worksheets for consistent data representation.
// AI Prompts: Write C# code that sets a List<Person> as the data source for WorkbookDesigner and uses smart markers on two worksheets. | Explain the syntax for accessing the first element of a collection in an Aspose.Cells variable smart marker. | Show the steps to process smart markers and save the workbook when using an IEnumerable data source. | Provide troubleshooting tips if smart markers are not populated after binding an IEnumerable to WorkbookDesigner.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple POCO class that will be used as the data source
    // Shows how to use a List<Person> (IEnumerable) as the data source for Aspose.Cells WorkbookDesigner, place smart markers on two sheets, process them, and save the workbook.
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
                // 1. Prepare sample data as an IEnumerable (List implements IEnumerable)
                List<Person> people = new List<Person>
                {
                    new Person("Alice", 30),
                    new Person("Bob", 45),
                    new Person("Charlie", 28)
                };

                // 2. Create a new workbook and add two worksheets
                Workbook workbook = new Workbook();

                // First worksheet (default sheet)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Employees";

                // Add second worksheet and obtain the Worksheet object directly
                Worksheet sheet2 = workbook.Worksheets.Add("Summary");

                // 3. Insert smart markers into the worksheets
                // Sheet1 will list each person's Name and Age
                sheet1.Cells["A1"].PutValue("Name");
                sheet1.Cells["B1"].PutValue("Age");
                sheet1.Cells["A2"].PutValue("&=Person.Name");
                sheet1.Cells["B2"].PutValue("&=Person.Age");

                // Sheet2 will display the first person's name using a variable marker
                sheet2.Cells["A1"].PutValue("First Employee:");
                sheet2.Cells["B1"].PutValue("&=$Person[0].Name"); // Access first element directly

                // 4. Create a WorkbookDesigner and bind the IEnumerable collection
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Person", people);

                // 5. Process the smart markers
                designer.Process();

                // 6. Save the resulting workbook
                string outputPath = "SmartMarker_IEnumerable_Output.xlsx";
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
