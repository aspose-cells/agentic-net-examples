// Title: Freeze specific rows and columns in an Excel worksheet by reading row/column indices from an XML config file using Aspose.Cells for .NET
// AI Prompts: Parse the <Row> and <Column> elements from a config.xml file and invoke Worksheet.FreezePanes with those values in a C# Aspose.Cells project. | Extend the sample to iterate over multiple worksheets, each with its own <Row> and <Column> settings defined under separate <Worksheet> nodes in the XML configuration. | Add validation that ensures the XML values are positive integers, provides default freeze positions when they are missing, and logs any parsing errors.
// Common Searches: C# Aspose.Cells read freeze pane coordinates from XML file | How to programmatically set freeze panes in Excel using Aspose.Cells and an external config | Load row and column numbers from config.xml and apply FreezePanes in .NET | Example of using XDocument to configure Excel freeze rows with Aspose.Cells
// Tags: Aspose.Cells FreezePanes from XML | C# read Excel freeze settings XML | Worksheet.FreezePanes with dynamic indices | XML configuration for Excel pane freezing | error handling missing XML elements Aspose.Cells

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The code loads an existing workbook, extracts <Row> and <Column> values from a config.xml file, applies Worksheet.FreezePanes using those indices on the first worksheet, and saves the updated workbook.
class FreezePaneFromXml
{
    static void Main()
    {
        // Paths to the workbook, XML configuration and output file
        string workbookPath = "input.xlsx";
        string configPath = "config.xml";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input files exist
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");
            if (!File.Exists(configPath))
                throw new FileNotFoundException($"Configuration file not found: {configPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Load the XML configuration file
            XDocument config = XDocument.Load(configPath);

            // Extract freeze row and column values from XML with safety checks
            XElement? rowElement = config.Root?.Element("Row");
            XElement? columnElement = config.Root?.Element("Column");

            if (rowElement == null || columnElement == null)
                throw new InvalidDataException("XML configuration must contain <Row> and <Column> elements.");

            int freezeRow = (int)rowElement;
            int freezeColumn = (int)columnElement;

            // Apply freeze panes: rows above freezeRow and columns left of freezeColumn are frozen
            // Worksheet.FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
