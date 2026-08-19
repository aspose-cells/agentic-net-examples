// Title: Aspose.Cells C# Smart Marker to Retrieve First Contact Phone Number (contacts[0].number)
// Description: Demonstrates how to place a smart marker "&=contacts[0].number" in cell A1, bind a List&lt;Contact&gt; as the data source named "contacts", process the marker with WorkbookDesigner, and save the populated Excel file.
// Keywords: Aspose.Cells smart marker | C# Excel smart marker list index | contacts[0].number example | WorkbookDesigner data source List<T> | populate Excel cell from collection | Aspose.Cells tutorial USA | Aspose.Cells GitHub sample
// Common Searches: Aspose.Cells reference first item in list smart marker | C# smart marker syntax contacts[0].number | How to bind List<Contact> to Aspose.Cells template | Insert phone number using smart marker in Excel | Aspose.Cells smart marker list indexing tutorial
// Developer Intent: Add a smart marker that pulls the phone number from the first Contact object in a list and generate the Excel workbook.
// Use Cases: Create a contact sheet that shows a single customer's phone number using a smart marker. | Generate an invoice where the primary client phone number is inserted from a List<Contact> data source. | Build a quick lookup report that displays the first entry of any collection via a smart marker.
// AI Prompts: Show how to modify the code to iterate over all contacts and write each phone number to successive rows using smart markers. | Provide an example that includes additional fields (e.g., name, email) in the Contact class and uses smart markers for each column. | Explain how to safely handle null or empty phone numbers when using the contacts[0].number smart marker.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerExample
{
    // Simple data class representing a contact with a phone number
    // Demonstrates how to place a smart marker "&=contacts[0].number" in cell A1, bind a List&lt;Contact&gt; as the data source named "contacts", process the marker with WorkbookDesigner, and save the populated Excel file.
    public class Contact
    {
        public string number { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that references the first contact's phone number
            // The syntax contacts[0].number accesses the first element of the contacts list
            sheet.Cells["A1"].PutValue("&=contacts[0].number");

            // Prepare the data source: a list containing a single Contact object
            List<Contact> contacts = new List<Contact>
            {
                new Contact { number = "555-1234" }
            };

            // Initialize WorkbookDesigner with the template workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Set the data source for the smart marker
            designer.SetDataSource("contacts", contacts);

            // Process the smart markers and populate the cell with the actual data
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("SmartMarkerNestedListOutput.xlsx");
        }
    }
}
