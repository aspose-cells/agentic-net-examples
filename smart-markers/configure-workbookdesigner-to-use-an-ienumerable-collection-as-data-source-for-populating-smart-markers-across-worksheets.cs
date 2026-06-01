using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Sample data class
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
                // Prepare an IEnumerable collection as data source
                List<Person> people = new List<Person>
                {
                    new Person("Alice", 30),
                    new Person("Bob", 25),
                    new Person("Charlie", 35)
                };

                // Create a new workbook (template) and add smart markers
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Employees";

                // Header row
                sheet1.Cells["A1"].PutValue("Name");
                sheet1.Cells["B1"].PutValue("Age");
                // Smart marker rows (the designer will repeat these rows for each collection item)
                sheet1.Cells["A2"].PutValue("&=Person.Name");
                sheet1.Cells["B2"].PutValue("&=Person.Age");

                // Add a second worksheet with its own smart markers
                int sheet2Index = workbook.Worksheets.Add();               // Returns the index of the new sheet
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Summary";
                sheet2.Cells["A1"].PutValue("Total Employees:");
                sheet2.Cells["B1"].PutValue("&=Person.Count"); // Count of the collection

                // Initialize WorkbookDesigner with the workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Bind the IEnumerable collection to the smart marker variable "Person"
                designer.SetDataSource("Person", people);

                // Process all smart markers in the workbook
                designer.Process();

                // Save the populated workbook
                string outputPath = "SmartMarker_Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}