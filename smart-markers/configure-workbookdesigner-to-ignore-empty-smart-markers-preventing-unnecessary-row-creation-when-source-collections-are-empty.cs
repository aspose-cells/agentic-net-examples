// Title: Aspose.Cells .NET: Use WorkbookDesigner.LineByLine = false to skip empty smart markers and avoid extra rows
// Description: Shows how to set WorkbookDesigner.LineByLine to false in C# so that smart markers defined in a named range are processed as a block. When the data source (e.g., List<Employee>) is empty, no placeholder rows are added, producing a clean workbook.
// Keywords: Aspose.Cells | WorkbookDesigner | LineByLine false | smart markers | ignore empty collection | prevent extra rows | C# .NET | named range | template processing | empty data source
// Common Searches: Aspose.Cells ignore empty smart markers | WorkbookDesigner LineByLine property effect | prevent blank rows in smart marker report | C# Aspose smart markers empty collection | how to skip rows when smart marker list is empty
// Developer Intent: Configure WorkbookDesigner so that an empty smart‑marker collection does not generate additional rows in the output workbook.
// Use Cases: Generate a spreadsheet report where a section may have zero records without leaving empty rows. | Create optional table blocks in a template that disappear when their data source is empty. | Process smart markers as a single block to maintain layout integrity when no data is supplied.
// AI Prompts: Provide C# code that sets WorkbookDesigner.LineByLine = false to ignore empty smart markers. | Explain how named ranges and LineByLine affect smart marker processing in Aspose.Cells. | Show an example of preventing extra rows when the smart marker data source is an empty list.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample data class
    // Shows how to set WorkbookDesigner.LineByLine to false in C# so that smart markers defined in a named range are processed as a block. When the data source (e.g., List<Employee>) is empty, no placeholder rows are added, producing a clean workbook.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class IgnoreEmptySmartMarkers
    {
        public static void Run()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Set up smart markers in the template
            // Header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            // Data row with smart markers
            sheet.Cells["A2"].PutValue("&=Employees.Name");
            sheet.Cells["B2"].PutValue("&=Employees.Age");

            // 3. Define a named range that contains the smart markers.
            // When LineByLine is false, the designer processes the range as a block.
            sheet.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // 4. Configure the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Disable line‑by‑line processing so that an empty collection does not create rows.
                LineByLine = false
            };

            // 5. Provide an empty collection as the data source.
            List<Employee> emptyEmployees = new List<Employee>(); // No items
            designer.SetDataSource("Employees", emptyEmployees);

            // 6. Process the smart markers.
            designer.Process();

            // 7. Save the result – no extra rows will be added because the collection is empty.
            workbook.Save("Result_IgnoringEmptySmartMarkers.xlsx");
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
