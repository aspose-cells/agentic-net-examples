// Title: Group nested parent‑child data with Aspose.Cells WorkbookDesigner – set LineByLine = false
// Description: Demonstrates how to bind a List<Parent> containing child collections, name the child range "_CellsSmartMarkers", set WorkbookDesigner.LineByLine to false, and process smart markers so the child rows are merged as a single grouped block before saving the Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | LineByLine | smart markers | nested data | parent child | grouped records | C# | Excel export | hierarchical data
// Common Searches: Aspose.Cells LineByLine false example | smart markers group child rows | bind nested collection Aspose.Cells C# | disable line‑by‑line processing Aspose.Cells | group parent child data in Excel using Aspose
// Developer Intent: Disable line‑by‑line processing so a child collection is merged into one smart‑marker block when using WorkbookDesigner.
// Use Cases: Generate an order report where each order appears once and its items are listed together in a grouped block. | Create an invoice sheet with a customer header and a single smart‑marker range for all line‑item rows. | Export categories and sub‑categories to Excel while keeping each category’s sub‑items in the same block. | Produce a multi‑level inventory sheet that groups product variants under their parent SKU.
// AI Prompts: Add a total quantity column for each parent after the grouped child rows while keeping LineByLine set to false. | Provide C# code that uses multiple "_CellsSmartMarkers" ranges to process several nested collections in one workbook. | Explain how to troubleshoot cases where smart markers do not group correctly after setting WorkbookDesigner.LineByLine to false.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    // Sample data classes representing nested (parent‑child) objects
    // Demonstrates how to bind a List<Parent> containing child collections, name the child range "_CellsSmartMarkers", set WorkbookDesigner.LineByLine to false, and process smart markers so the child rows are merged as a single grouped block before saving the Excel file.
    public class Parent
    {
        public string Name { get; set; }
        public List<Child> Children { get; set; }
    }

    public class Child
    {
        public string Item { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (template) and add smart markers.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Smart markers for parent record
            ws.Cells["A1"].PutValue("&Parent.Name");

            // Smart markers for child collection – placed in a range that will be processed as a group.
            // The range name "_CellsSmartMarkers" tells the designer to treat the whole block as one group
            // when LineByLine is set to false.
            ws.Cells["A2"].PutValue("&Children.Item");
            ws.Cells["B2"].PutValue("&Children.Quantity");
            ws.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";

            // 2. Prepare nested data.
            var data = new List<Parent>
            {
                new Parent
                {
                    Name = "Order001",
                    Children = new List<Child>
                    {
                        new Child { Item = "Apple",  Quantity = 10 },
                        new Child { Item = "Banana", Quantity = 5 }
                    }
                },
                new Parent
                {
                    Name = "Order002",
                    Children = new List<Child>
                    {
                        new Child { Item = "Orange", Quantity = 8 },
                        new Child { Item = "Grape",  Quantity = 12 }
                    }
                }
            };

            // 3. Initialize WorkbookDesigner, assign the workbook, and set LineByLine to false.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb,
                LineByLine = false   // Process nested objects as grouped records
            };

            // 4. Bind the data source. The root name "Parent" matches the smart marker prefix.
            designer.SetDataSource("Parent", data);

            // 5. Process the smart markers.
            designer.Process();

            // 6. Save the resulting workbook.
            wb.Save("NestedObjectsGrouped.xlsx");
        }
    }
}
