using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerProtection
{
    public class Program
    {
        public static void Main()
        {
            // Load the workbook template that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare a simple data source for the smart markers
            var people = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 28 }
            };

            // Set the data source and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("People", people);
            designer.Process();

            // Unlock cells that should remain editable after the worksheet is protected
            Worksheet sheet = workbook.Worksheets[0];
            // Example: unlock column B (index 1) for rows 2 and 3 (zero‑based indices 1 and 2)
            for (int row = 1; row <= 2; row++)
            {
                Cell cell = sheet.Cells[row, 1];
                Style style = cell.GetStyle();
                style.IsLocked = false; // make this cell editable when the sheet is protected
                cell.SetStyle(style);
            }

            // Configure protection options
            Protection protection = sheet.Protection;
            protection.AllowSelectingUnlockedCell = true; // allow users to select unlocked cells
            protection.Password = "pwd123";               // set a password for the worksheet

            // Protect the worksheet with all protection types and the specified password
            sheet.Protect(ProtectionType.All, protection.Password, null);

            // Save the processed and protected workbook
            workbook.Save("SmartMarkersProtected.xlsx");
        }

        // Simple POCO class used as a data source for smart markers
        public class Person
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }
    }
}