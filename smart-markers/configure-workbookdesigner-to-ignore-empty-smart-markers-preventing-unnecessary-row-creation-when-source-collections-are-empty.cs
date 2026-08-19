// Title: Aspose.Cells C# – Skip Empty Smart Markers Using WorkbookDesigner.LineByLine = false
// Description: Demonstrates how to configure WorkbookDesigner to ignore empty smart‑marker collections. The example creates a workbook, defines range smart markers for an "Employees" list, sets LineByLine to false, supplies an empty List<Employee>, processes the markers, and saves the file without adding extra rows.
// Keywords: Aspose.Cells | C# | .NET | WorkbookDesigner | LineByLine false | ignore empty smart markers | range smart markers | smart marker collection empty | prevent row insertion | Excel report template
// Common Searches: Aspose.Cells skip empty smart markers | WorkbookDesigner LineByLine property example | prevent rows from being added when smart marker list is empty | range smart markers with empty data source | C# Aspose.Cells ignore empty collection
// Developer Intent: Configure WorkbookDesigner so that an empty smart‑marker collection does not generate additional rows in the output workbook.
// Use Cases: Create a reusable Excel template where a section may have zero records, keeping only the header row. | Apply named range smart markers (_CellsSmartMarkers) and preserve worksheet layout when the source list is empty. | Combine LineByLine = false with an empty data source to maintain formatting while still processing other markers.
// AI Prompts: Show how to set WorkbookDesigner.LineByLine = false to ignore empty smart markers in Aspose.Cells C#. | Provide a C# code snippet that defines a named range for smart markers and processes an empty collection without adding rows. | Explain the effect of the LineByLine property on range smart markers when the data source collection is empty.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample data class
    // Demonstrates how to configure WorkbookDesigner to ignore empty smart‑marker collections. The example creates a workbook, defines range smart markers for an "Employees" list, sets LineByLine to false, supplies an empty List<Employee>, processes the markers, and saves the file without adding extra rows.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class IgnoreEmptySmartMarkers
    {
        public static void Run()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Set up smart markers for a collection named "Employees"
            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            // Data row (smart markers)
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // 4. Define a named range that contains the smart markers.
            // This tells the designer to work with range smart markers.
            sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // 5. Configure the designer to process range smart markers (LineByLine = false)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = false // important: disables line‑by‑line processing
            };

            // 6. Provide an empty collection as the data source.
            // No rows will be created because the collection is empty.
            List<Employee> emptyEmployees = new List<Employee>();
            designer.SetDataSource("Employees", emptyEmployees);

            // 7. Process the smart markers.
            designer.Process();

            // 8. Save the result (lifecycle: save)
            workbook.Save("IgnoreEmptySmartMarkers_Output.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            IgnoreEmptySmartMarkers.Run();
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}
