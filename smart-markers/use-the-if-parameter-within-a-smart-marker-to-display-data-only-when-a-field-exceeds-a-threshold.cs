// Title: C# – Use IF parameter in Aspose.Cells smart marker to display values above a threshold
// Description: Demonstrates how to create a workbook, add a header, and place a smart marker that uses the IF parameter to output the $Value field only when it exceeds 50. A list of Record objects is bound as the data source, processed with WorkbookDesigner, and saved as SmartMarkerIfDemo.xlsx.
// Keywords: Aspose.Cells IF smart marker | C# conditional smart marker | Excel threshold example Aspose | WorkbookDesigner conditional display | smart marker numeric filter
// Common Searches: Aspose.Cells IF smart marker example C# | show values greater than 50 using smart markers | conditional smart marker syntax Aspose | filter rows with smart markers in .NET | how to use IF parameter in Aspose.Cells
// Developer Intent: Implement a smart marker that writes a cell value only when the source field meets a numeric condition, using the IF parameter in Aspose.Cells for C#.
// Use Cases: Generate a sales report that lists only transactions above a target amount. | Create an inventory sheet that displays items with stock levels exceeding a reorder point. | Build a KPI dashboard that shows metrics only when they surpass a defined benchmark.
// AI Prompts: Provide C# code using Aspose.Cells smart markers with an IF parameter to show values greater than 100. | Explain how to combine multiple IF conditions in smart markers to apply different formatting based on numeric ranges. | Show how to bind a collection of objects to a smart marker and filter the output with conditional logic.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    // Demonstrates how to create a workbook, add a header, and place a smart marker that uses the IF parameter to output the $Value field only when it exceeds 50. A list of Record objects is bound as the data source, processed with WorkbookDesigner, and saved as SmartMarkerIfDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header cell
            sheet.Cells["A1"].PutValue("Value");

            // Smart marker with IF parameter:
            // Displays the value only when it is greater than 50
            sheet.Cells["A2"].PutValue("&IF($Value>50,$Value,\"\")");

            // Prepare a data source
            List<Record> records = new List<Record>
            {
                new Record { Value = 30 },
                new Record { Value = 60 },
                new Record { Value = 45 },
                new Record { Value = 80 }
            };

            // Set up the designer, assign the data source and process the smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", records);
            designer.Process();

            // Save the resulting workbook
            workbook.Save("SmartMarkerIfDemo.xlsx");
        }

        // Simple data class used as the data source
        public class Record
        {
            public int Value { get; set; }
        }
    }
}
