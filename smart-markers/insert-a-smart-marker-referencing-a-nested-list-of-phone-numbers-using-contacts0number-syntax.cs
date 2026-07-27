using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerExample
{
    // Simple data model: a contact with a list of phone numbers
    public class Contact
    {
        public List<string> number { get; set; } = new List<string>();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Insert a smart marker that references the first contact's phone numbers
            //    The syntax contacts[0].number tells the designer to iterate over the nested list.
            sheet.Cells["A1"].PutValue("&=contacts[0].number");

            // 3. Prepare sample data: a list of contacts, each containing a list of phone numbers
            List<Contact> contacts = new List<Contact>();

            Contact contact1 = new Contact();
            contact1.number.Add("555-0101");
            contact1.number.Add("555-0102");
            contacts.Add(contact1);

            Contact contact2 = new Contact();
            contact2.number.Add("555-0201");
            contacts.Add(contact2);

            // 4. Set up the WorkbookDesigner, assign the data source, and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("contacts", contacts);
            designer.Process(); // lifecycle rule: process smart markers

            // 5. Save the result (lifecycle rule: save)
            workbook.Save("SmartMarker_NestedList_Output.xlsx");
        }
    }
}