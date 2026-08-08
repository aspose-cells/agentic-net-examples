// Title: C# – Load Excel Template and Apply CopyStyle Smart Marker to Preserve Formatting with Aspose.Cells
// Description: Loads a workbook template, binds a List<Record> to a smart‑marker named Data, processes the template with WorkbookDesigner, and saves the output. The CopyStyle attribute on the markers makes every generated row inherit the original cell styles (fonts, borders, number formats).
// Keywords: Aspose.Cells | CopyStyle | smart markers | C# | WorkbookDesigner | Excel template | preserve formatting | populate rows | automation
// Common Searches: Aspose.Cells CopyStyle example C# | how to keep cell formatting when adding rows with smart markers | load Excel template and preserve styles using Aspose.Cells | WorkbookDesigner SetDataSource with CopyStyle attribute | smart marker copy style .NET
// Developer Intent: Generate rows from a data collection while automatically inheriting the template’s cell formatting via the CopyStyle smart marker.
// Use Cases: Create invoices where each line‑item row retains the header’s font, border, and number format. | Build sales or budget reports from a list of objects without losing conditional formatting defined in the template. | Export dynamic data to a pre‑styled Excel sheet for downstream processing or printing.
// AI Prompts: Add a total row that also uses the CopyStyle attribute to match the existing style. | Explain how the CopyStyle attribute works with smart markers in Aspose.Cells for .NET, step by step. | Generate C# code that reads data from a DataTable and applies CopyStyle smart markers to a workbook template.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCopyStyleDemo
{
    // Sample data class representing a record
    // Loads a workbook template, binds a List<Record> to a smart‑marker named Data, processes the template with WorkbookDesigner, and saves the output. The CopyStyle attribute on the markers makes every generated row inherit the original cell styles (fonts, borders, number formats).
    public class Record
    {
        public string Name { get; set; }
        public double Amount { get; set; }

        public Record(string name, double amount)
        {
            Name = name;
            Amount = amount;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the workbook template that contains smart markers with the CopyStyle attribute
            string templatePath = "Template.xlsx";

            // Load the template workbook
            WorkbookDesigner designer = new WorkbookDesigner();
            designer.Workbook = new Workbook(templatePath);

            // Prepare sample data source
            List<Record> records = new List<Record>
            {
                new Record("Alice", 123.45),
                new Record("Bob", 678.90),
                new Record("Charlie", 234.56)
            };

            // Set the data source for the smart markers (assumes markers like &=CopyStyle&=Data.Name, &Data.Amount)
            designer.SetDataSource("Data", records);

            // Process the template – this will generate rows and inherit formatting via the CopyStyle attribute
            designer.Process();

            // Save the resulting workbook
            designer.Workbook.Save("Output.xlsx");
        }
    }
}
