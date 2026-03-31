using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomDataSource
{
    // Sample class representing a custom data source
    public class Level
    {
        public string Name { get; set; }
        public string Value { get; set; }

        public Level(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers (e.g., &Level.Name, &Level.Value)
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Create a custom data source (a list of Level objects)
            List<Level> levels = new List<Level>
            {
                new Level("Row1", "Value1"),
                new Level("Row2", "Value2"),
                new Level("Row3", "Value3")
            };

            // Bind the custom data source to the smart marker name "Level"
            designer.SetDataSource("Level", levels);

            // Process the smart markers and populate the worksheet with data from the custom source
            designer.Process();

            // Save the populated workbook
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}