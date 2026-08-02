// Title: Load Freeze Pane Settings from XML and Apply with Aspose.Cells (C#)
// Description: A C# example that reads row, column, frozenRows and frozenColumns values from an XML file, creates a workbook with sample data, applies Worksheet.FreezePanes using Aspose.Cells, and saves the result as Result.xlsx.
// Keywords: Aspose.Cells | C# | FreezePanes | XML configuration | worksheet freeze | Excel automation | read XML settings | dynamic freeze rows | dynamic freeze columns
// Common Searches: Aspose.Cells read freeze pane values from XML | C# set FreezePanes using external config | load freeze rows and columns from file Aspose.Cells | apply XML based freeze pane settings in Excel | dynamic freeze pane example C# Aspose
// Developer Intent: Read freeze‑pane coordinates from an XML file and apply them to a worksheet with Aspose.Cells.
// Use Cases: Generate reports where the frozen area is defined by a user‑editable XML template. | Standardize navigation across multiple workbooks by applying a common freeze configuration. | Create a batch process that reads different XML files to set unique freeze panes for each worksheet.
// AI Prompts: Write C# code that reads freeze pane parameters from a JSON file and uses Aspose.Cells to apply them. | Show how to validate XML elements and provide fallback values when setting FreezePanes in Aspose.Cells. | Demonstrate applying separate freeze pane settings to several worksheets using a single XML configuration.

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// A C# example that reads row, column, frozenRows and frozenColumns values from an XML file, creates a workbook with sample data, applies Worksheet.FreezePanes using Aspose.Cells, and saves the result as Result.xlsx.
class FreezeFromConfig
{
    static void Main()
    {
        // Path to the XML configuration file that contains freeze pane settings
        string configPath = "freezeConfig.xml";

        // Default values (no freeze) – will be overwritten if the file exists and contains data
        int row = 0;
        int column = 0;
        int frozenRows = 0;
        int frozenColumns = 0;

        // Load configuration values from XML
        if (File.Exists(configPath))
        {
            XDocument doc = XDocument.Load(configPath);
            XElement root = doc.Element("FreezeConfig");
            if (root != null)
            {
                // Use nullable cast to avoid exceptions if an element is missing
                row = (int?)root.Element("Row") ?? 0;
                column = (int?)root.Element("Column") ?? 0;
                frozenRows = (int?)root.Element("FrozenRows") ?? 0;
                frozenColumns = (int?)root.Element("FrozenColumns") ?? 0;
            }
        }

        // Create a new workbook (Aspose.Cells)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Optional: populate some sample data so the effect of freezing can be seen
        for (int i = 0; i < 20; i++)
        {
            for (int j = 0; j < 5; j++)
            {
                worksheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Apply freeze panes if the configuration specifies a non‑zero position
        // FreezePanes(row, column, frozenRows, frozenColumns) uses zero‑based indices
        if (row > 0 || column > 0)
        {
            worksheet.FreezePanes(row, column, frozenRows, frozenColumns);
        }

        // Save the workbook to an Excel file
        workbook.Save("Result.xlsx");
    }
}
