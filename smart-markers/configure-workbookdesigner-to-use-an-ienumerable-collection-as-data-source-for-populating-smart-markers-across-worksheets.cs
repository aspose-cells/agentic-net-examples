// Title: Bind an IEnumerable<Person> collection to smart markers on multiple worksheets using WorkbookDesigner in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a List<Person>, adds smart marker cells to two worksheets, binds the list to the 'Person' variable with WorkbookDesigner, processes all markers, and saves the workbook as an .xlsx file. | Demonstrate how to configure WorkbookDesigner to accept any IEnumerable data source for smart markers and apply it to every sheet in an Aspose.Cells workbook.
// Common Searches: asp.net bind IEnumerable to smart markers with WorkbookDesigner Aspose.Cells | populate smart markers on several worksheets from a List<T> in C# | how to set data source for smart markers using a collection in Aspose.Cells .NET | example of processing smart markers across multiple sheets in Aspose.Cells | using WorkbookDesigner to generate Excel reports from an IEnumerable collection
// Tags: WorkbookDesigner bind IEnumerable data source | smart markers populate multiple worksheets | Aspose.Cells set data source from List<Person> | C# process smart markers across sheets | Excel template generation with smart markers

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple POCO class that will be used as the data source
    // The example creates a workbook with two worksheets containing smart markers, binds a List<Person> (IEnumerable) to the 'Person' variable via WorkbookDesigner, processes all markers across the sheets, and saves the populated Excel file.
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

    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare sample data as an IEnumerable (List<Person>)
                List<Person> persons = new List<Person>
                {
                    new Person("Alice", 30),
                    new Person("Bob", 45),
                    new Person("Charlie", 28)
                };

                // Create a new workbook that will act as the template
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 1 – simple smart markers for Name & Age
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Employees";

                // Header row
                sheet1.Cells["A1"].PutValue("Name");
                sheet1.Cells["B1"].PutValue("Age");

                // Smart markers – the designer will repeat these rows for each Person
                sheet1.Cells["A2"].PutValue("&=Person.Name");
                sheet1.Cells["B2"].PutValue("&=Person.Age");

                // -------------------------------------------------
                // Worksheet 2 – another set of smart markers
                // -------------------------------------------------
                int newSheetIndex = workbook.Worksheets.Add();               // Add returns the index of the new sheet
                Worksheet sheet2 = workbook.Worksheets[newSheetIndex];
                sheet2.Name = "Summary";

                // Header row
                sheet2.Cells["A1"].PutValue("Employee");
                sheet2.Cells["B1"].PutValue("Years");

                // Smart markers referencing the same data source
                sheet2.Cells["A2"].PutValue("&=Person.Name");
                sheet2.Cells["B2"].PutValue("&=Person.Age");

                // -------------------------------------------------
                // Configure WorkbookDesigner
                // -------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Bind the IEnumerable collection to the smart marker variable "Person"
                designer.SetDataSource("Person", persons);

                // Process all smart markers in the workbook
                designer.Process();

                // Save the populated workbook
                string outputPath = "SmartMarkersFromIEnumerable.xlsx";
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
