// Title: Apply Freeze Panes from XML Config with Aspose.Cells for .NET (C#)
// Description: C# example that loads an XML file (freezeConfig.xml), reads the <FreezeRow> and <FreezeColumn> elements (default 0), creates a workbook, applies Worksheet.FreezePanes using those values, and saves the result as FreezePanesResult.xlsx.
// Keywords: Aspose.Cells | C# | XML configuration | freeze panes | FreezePanes method | read XML values | dynamic row freeze | Excel automation | workbook save | programmatic layout
// Common Searches: Aspose.Cells read freeze row from XML | C# set freeze panes using XML config | How to freeze panes programmatically with Aspose.Cells | Load freeze pane settings from XML in .NET | XML driven freeze panes Aspose.Cells example
// Developer Intent: Read freeze‑row/column values from an XML file and apply them as freeze panes to a worksheet using Aspose.Cells.
// Use Cases: Generate Excel reports where the freeze pane position is defined by an external XML file, allowing end‑users to control view layout without code changes. | Apply freeze panes only when the configuration specifies non‑zero rows or columns, keeping the sheet unfrozen otherwise. | Integrate the XML‑based freeze logic into an automated reporting pipeline that creates workbooks with consistent scrolling behavior across multiple environments.
// AI Prompts: Write C# code that reads <FreezeRow> and <FreezeColumn> from an XML file and uses Aspose.Cells to set FreezePanes on the first worksheet. | Show how to handle missing or malformed XML configuration when applying freeze panes with Aspose.Cells, including default fallback values and error logging. | Extend the example to support multiple worksheets, each with its own freeze settings defined in the same XML configuration.

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

namespace Example
{
    // C# example that loads an XML file (freezeConfig.xml), reads the <FreezeRow> and <FreezeColumn> elements (default 0), creates a workbook, applies Worksheet.FreezePanes using those values, and saves the result as FreezePanesResult.xlsx.
    class FreezePanesFromConfig
    {
        static void Main()
        {
            try
            {
                // Path to the XML configuration file
                string configPath = "freezeConfig.xml";

                // Default freeze values
                int freezeRow = 0;
                int freezeColumn = 0;

                // Load configuration if the file exists
                if (File.Exists(configPath))
                {
                    XDocument configDoc = XDocument.Load(configPath);
                    freezeRow = (int?)configDoc.Root.Element("FreezeRow") ?? 0;
                    freezeColumn = (int?)configDoc.Root.Element("FreezeColumn") ?? 0;
                }
                else
                {
                    // Configuration file not found; proceeding with defaults (no freeze)
                }

                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Apply freeze panes if needed
                if (freezeRow > 0 || freezeColumn > 0)
                {
                    sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);
                }

                // Save the workbook
                string outputPath = "FreezePanesResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
