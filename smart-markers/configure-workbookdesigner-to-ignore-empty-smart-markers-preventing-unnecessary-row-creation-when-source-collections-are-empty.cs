// Title: How to configure Aspose.Cells WorkbookDesigner to ignore empty smart marker collections and prevent extra rows in C#
// AI Prompts: Write C# code that sets WorkbookDesigner.LineByLine = false, binds an empty List<T> to a smart marker, and processes the template without inserting rows. | Show an example of using a named range for smart markers in Aspose.Cells to avoid row creation when the data source collection is empty. | Demonstrate how to prevent Aspose.Cells from adding placeholder rows for empty smart marker collections in a .NET workbook.
// Common Searches: Aspose.Cells C# ignore empty smart marker collection | prevent row insertion when smart marker data source is empty in .NET | WorkbookDesigner LineByLine false usage example | named range smart markers avoid extra rows Aspose.Cells | skip smart marker processing for empty lists in C#
// Tags: WorkbookDesigner line-by-line mode disabled | smart marker named range usage | avoid row creation for empty collections | Aspose.Cells smart marker empty list handling | C# workbook designer empty data source

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data class for demonstration
    // The example creates a workbook with a header row and a smart‑marker row, defines a named range for the markers, binds an empty List<Employee> to the marker name, disables line‑by‑line processing by setting WorkbookDesigner.LineByLine to false, processes the smart markers, and saves the file so that only the header remains and no extra rows are added.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class IgnoreEmptySmartMarkersExample
    {
        public static void Run()
        {
            // 1. Create a new workbook (template) and add smart markers
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            // Smart marker row – will be repeated for each item in the collection
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // Define a named range that contains the smart markers.
            // When LineByLine is false, the designer processes only this range.
            sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // 2. Prepare an empty data source (no rows should be added)
            List<Employee> emptyList = new List<Employee>(); // empty collection

            // 3. Configure WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Disable line‑by‑line processing so that the designer works on the named range.
                // When the collection is empty, no rows are inserted.
                LineByLine = false
            };

            // Bind the empty collection to the smart marker name "Employees"
            designer.SetDataSource("Employees", emptyList);

            // 4. Process the smart markers
            designer.Process();

            // 5. Save the result – the worksheet will contain only the header row,
            // no extra rows created for the empty collection.
            workbook.Save("IgnoreEmptySmartMarkersOutput.xlsx");
        }
    }

    // Entry point for testing
    class Program
    {
        static void Main()
        {
            IgnoreEmptySmartMarkersExample.Run();
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
