// Title: Create an Excel workbook from nested Parent‑Child objects using Aspose.Cells smart marker range in C#
// AI Prompts: Generate a C# program that defines smart markers for a parent collection and its child collection, then uses WorkbookDesigner to fill an Excel sheet with department and employee names. | Demonstrate how to flatten the child list, bind it as a separate data source, assign the name "_CellsSmartMarkers" to the marker range, and call Process to generate the final workbook. | Write code that saves the resulting workbook to a file, handles any exceptions, and prints a confirmation message.
// Common Searches: Aspose.Cells smart markers nested collections example C# | how to bind parent and child lists to WorkbookDesigner for Excel export | named smart marker range processing with Aspose.Cells | flatten child objects for smart markers Aspose.Cells tutorial | generate Excel file from hierarchical data using smart markers C#
// Tags: Aspose.Cells WorkbookDesigner bind parent collection | smart marker range naming Aspose.Cells | populate worksheet with hierarchical data C# | process nested collections using smart markers | export object hierarchy to Excel with Aspose.Cells

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerNestedExample
{
    // Sample parent class containing a collection of children
    // The example creates a workbook, places smart markers "&=Parents.Name" and "&=Parents.Children.Name" in a named range, builds a list of Parent objects each with a list of Child objects, flattens the children into a separate collection, binds both collections to a WorkbookDesigner, processes the smart markers to repeat rows for each department and its employees, and saves the result as NestedSmartMarkersOutput.xlsx.
    public class Parent
    {
        public string? Name { get; set; }
        public List<Child>? Children { get; set; }
    }

    // Sample child class
    public class Child
    {
        public string? Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook(); // lifecycle: create
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Set up smart markers in the worksheet
                // Header row
                sheet.Cells["A1"].PutValue("Parent");
                sheet.Cells["B1"].PutValue("Child");

                // Data rows with smart markers
                // "&=Parents.Name" will repeat for each Parent item
                // "&=Parents.Children.Name" will repeat for each Child of the current Parent
                sheet.Cells["A2"].PutValue("&=Parents.Name");
                sheet.Cells["B2"].PutValue("&=Parents.Children.Name");

                // Define the range that contains the smart markers and give it the required name
                Aspose.Cells.Range smartRange = sheet.Cells.CreateRange("A2:B2");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Prepare nested data source
                List<Parent> parents = new List<Parent>
                {
                    new Parent
                    {
                        Name = "Department A",
                        Children = new List<Child>
                        {
                            new Child { Name = "Alice" },
                            new Child { Name = "Bob" }
                        }
                    },
                    new Parent
                    {
                        Name = "Department B",
                        Children = new List<Child>
                        {
                            new Child { Name = "Charlie" },
                            new Child { Name = "Diana" }
                        }
                    }
                };

                // Flatten child collection for separate data source (required by smart markers)
                List<Child> allChildren = parents.SelectMany(p => p.Children ?? new List<Child>()).ToList();

                // 4. Initialize WorkbookDesigner and bind data sources
                WorkbookDesigner designer = new WorkbookDesigner(); // lifecycle: create
                designer.Workbook = workbook;
                designer.SetDataSource("Parents", parents);
                designer.SetDataSource("Children", allChildren);

                // 5. Process the smart markers (range is already named, so parameterless Process is sufficient)
                designer.Process(); // lifecycle: process

                // 6. Save the resulting workbook
                string outputPath = "NestedSmartMarkersOutput.xlsx";
                workbook.Save(outputPath); // lifecycle: save

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
