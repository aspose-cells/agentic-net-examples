using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLineByLineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // ------------------------------------------------------------
            // Set up smart markers.
            // A root level marker for the parent object.
            ws.Cells["A1"].PutValue("&Root.Name");
            // Markers for the child collection. These will be processed as a group.
            ws.Cells["A2"].PutValue("&Root.Children.Name");
            ws.Cells["B2"].PutValue("&Root.Children.Age");

            // When LineByLine is false the designer expects a named range
            // that contains the markers for the grouped records.
            ws.Cells.CreateRange("A2:B2").Name = "_CellsSmartMarkers";
            // ------------------------------------------------------------

            // Prepare nested data: a parent object with a list of child objects.
            var root = new RootData
            {
                Name = "Company XYZ",
                Children = new List<ChildData>
                {
                    new ChildData { Name = "Alice",   Age = 30 },
                    new ChildData { Name = "Bob",     Age = 35 },
                    new ChildData { Name = "Charlie", Age = 28 }
                }
            };

            // Initialize the WorkbookDesigner, assign the workbook,
            // and set LineByLine to false so the child collection is processed as a group.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb,
                LineByLine = false
            };

            // Bind the data source to the smart marker name "Root".
            designer.SetDataSource("Root", root);

            // Process the smart markers and populate the worksheet.
            designer.Process();

            // Save the resulting workbook.
            wb.Save("LineByLineFalseOutput.xlsx");
        }

        // --------------------------------------------------------------------
        // Data classes used for the demonstration.
        // --------------------------------------------------------------------
        public class RootData
        {
            public string Name { get; set; }
            public List<ChildData> Children { get; set; }
        }

        public class ChildData
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }
    }
}