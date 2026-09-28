// Title: Insert a smart marker for a nested phone number list (contacts[0].number) using Aspose.Cells WorkbookDesigner in C#
// AI Prompts: Generate C# code that places the smart marker '&=contacts[0].number' in a worksheet cell, binds a List<Contact> where Contact contains a List<string> of phone numbers, and processes it with WorkbookDesigner. | Demonstrate how to use Aspose.Cells WorkbookDesigner to populate an Excel file from hierarchical data by inserting a smart marker that references the first contact's phone numbers.
// Common Searches: aspnet how to bind a list of objects with a list property to Aspose.Cells smart markers | example of using contacts[0].number smart marker in Aspose.Cells C# | populate Excel column with multiple phone numbers from nested collection using WorkbookDesigner | Aspose.Cells smart marker syntax for accessing nested list elements in C#
// Tags: Aspose.Cells WorkbookDesigner nested list | C# smart marker contacts[0].number | bind hierarchical data to Excel Aspose.Cells | populate Excel phone numbers from list of strings | smart marker list of strings Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerExample
{
    // Simple data model: each contact has a list of phone numbers called "number"
    // // This example creates a workbook, inserts the smart marker '&=contacts[0].number' into cell A1, binds a List<Contact> (each Contact holding a List<string> of phone numbers), processes the marker with WorkbookDesigner, and saves the populated file as SmartMarker_NestedList_Output.xlsx.
    public class Contact
    {
        public List<string> number { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert a smart marker that references the first contact's phone numbers list
            //    The syntax contacts[0].number will be resolved by the designer during processing
            sheet.Cells["A1"].PutValue("&=contacts[0].number");

            // 3. Prepare sample data: a list of contacts, each with a list of phone numbers
            List<Contact> contacts = new List<Contact>
            {
                new Contact
                {
                    number = new List<string> { "123-456-7890", "987-654-3210" }
                },
                new Contact
                {
                    number = new List<string> { "555-111-2222", "555-333-4444" }
                }
            };

            // 4. Set up the WorkbookDesigner, assign the workbook, and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = workbook;
            designer.SetDataSource("contacts", contacts);

            // 5. Process the smart markers to populate the worksheet with data
            designer.Process();

            // 6. Save the result
            workbook.Save("SmartMarker_NestedList_Output.xlsx");
        }
    }
}
