// Title: Aspose.Cells for .NET – Populate Merged Cells with Smart Markers and Auto‑Expand Rows
// Description: Demonstrates how to create a workbook, merge a header (A1:C1) and a data row (A2:C2), name the merged range as _CellsSmartMarkers, bind a List<Person> to smart markers (&=Data.Name, &=Data.Age), and use WorkbookDesigner to process the template so that the merged row repeats and expands for every record.
// Keywords: Aspose.Cells | smart markers | merged cells | auto expand rows | WorkbookDesigner | C# | .NET | named smart marker range | populate merged range | list data source
// Common Searches: Aspose.Cells smart markers merged cells example | auto expand merged rows with WorkbookDesigner | populate merged range from List<T> in C# | how to name smart marker range _CellsSmartMarkers | fill merged header and data rows using Aspose.Cells
// Developer Intent: Generate a spreadsheet where a merged block is filled via smart markers and automatically repeats for each item in a collection.
// Use Cases: Create an employee directory with a merged title row and a merged data row that expands for every employee object. | Design a financial report template where a merged section (e.g., A2:C2) serves as a repeating block for transaction records. | Build a printable invoice where product lines are placed in a merged area that grows with the number of line items.
// AI Prompts: Show code that defines a merged range as a smart marker, binds a List<Person>, and processes it so the merged rows expand automatically. | Explain why the merged cells must be created before calling WorkbookDesigner.Process and why the range name _CellsSmartMarkers is required. | Provide a step‑by‑step guide to set up a merged header, add smart markers in a merged data row, and save the workbook using Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkersMergedCellsDemo
{
    // Sample data class
    // Demonstrates how to create a workbook, merge a header (A1:C1) and a data row (A2:C2), name the merged range as _CellsSmartMarkers, bind a List<Person> to smart markers (&=Data.Name, &=Data.Age), and use WorkbookDesigner to process the template so that the merged row repeats and expands for every record.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ----- Template setup -----
                // Header row (merged across A1:C1)
                cells["A1"].PutValue("Employee List");
                cells.Merge(0, 0, 1, 3); // Merge A1:C1

                // Data row with smart markers (row 2 -> index 1)
                cells["A2"].PutValue("&=Data.Name"); // Smart marker for Name
                cells["B2"].PutValue("&=Data.Age");  // Smart marker for Age
                cells["C2"].PutValue("Static Text"); // Additional column (optional)

                // Merge the data row across A2:C2 to demonstrate expansion
                cells.Merge(1, 0, 1, 3); // Merge A2:C2

                // Define the smart marker range (required for processing)
                AsposeRange smartRange = cells.CreateRange("A2:C2");
                smartRange.Name = "_CellsSmartMarkers";

                // ----- Data source -----
                List<Person> persons = new List<Person>
                {
                    new Person { Name = "John Doe", Age = 30 },
                    new Person { Name = "Jane Smith", Age = 28 },
                    new Person { Name = "Bob Johnson", Age = 45 }
                };

                // ----- Process smart markers -----
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", persons);
                designer.Process(); // Populate data; merged rows will expand automatically

                // Save the result (lifecycle save)
                workbook.Save("SmartMarkersMergedCellsOutput.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
