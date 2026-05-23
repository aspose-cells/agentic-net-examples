using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsSmartMarkerExample
{
    // Simple data model: a contact with a list of phone numbers
    public class Contact
    {
        public string Name { get; set; }
        public List<string> Number { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (template)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Place a smart marker that references a nested list.
            //    The syntax contacts[0].Number will fetch the first phone number of the first contact.
            //    The marker must start with "&=" to be recognized by Aspose.Cells.
            sheet.Cells["A1"].PutValue("&=contacts[0].Number");

            // 3. Prepare sample data
            var contacts = new List<Contact>
            {
                new Contact
                {
                    Name = "John Doe",
                    Number = new List<string> { "123-456-7890", "098-765-4321" }
                },
                new Contact
                {
                    Name = "Jane Smith",
                    Number = new List<string> { "555-111-2222" }
                }
            };

            // 4. Initialize WorkbookDesigner, assign the workbook and set the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("contacts", contacts);

            // 5. Process the smart markers to populate the data
            designer.Process();

            // 6. Save the result
            workbook.Save("SmartMarker_NestedList_Output.xlsx");
        }
    }
}