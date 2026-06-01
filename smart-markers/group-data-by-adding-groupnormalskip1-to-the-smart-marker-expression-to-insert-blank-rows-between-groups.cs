using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables; // for WorkbookDesigner namespace

namespace AsposeCellsSmartMarkerGrouping
{
    // Simple data class for the smart marker data source
    public class Record
    {
        public string Category { get; set; }
        public double Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up header cells
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Insert smart markers with grouping instruction.
            // (group:normal,skip:1) inserts a blank row after each group.
            sheet.Cells["A2"].PutValue("&=Data.Group:normal,skip:1");
            sheet.Cells["B2"].PutValue("&=Data.Value");

            // Prepare sample data source
            List<Record> data = new List<Record>
            {
                new Record { Category = "Fruit", Value = 120 },
                new Record { Category = "Fruit", Value = 150 },
                new Record { Category = "Vegetable", Value = 80 },
                new Record { Category = "Vegetable", Value = 95 },
                new Record { Category = "Grain", Value = 60 }
            };

            // Initialize WorkbookDesigner with the workbook (lifecycle load)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the data source to the smart marker name "Data"
            designer.SetDataSource("Data", data);

            // Process the smart markers (uses the provided Process() rule)
            designer.Process();

            // Save the resulting workbook (lifecycle save)
            workbook.Save("GroupedSmartMarkers.xlsx");
        }
    }
}