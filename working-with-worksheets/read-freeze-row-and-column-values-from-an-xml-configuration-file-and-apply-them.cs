// Title: Apply freeze pane rows and columns to an Excel worksheet from XML settings using Aspose.Cells for .NET (C#)
// Common Searches: C# read freeze pane row and column values from config.xml and apply with Aspose.Cells | How to set Excel freeze panes in Aspose.Cells using external XML settings | Load freeze pane coordinates from an XML file and freeze first worksheet in .NET | Aspose.Cells FreezePanes method with parameters read from XML configuration | Programmatically freeze rows and columns in Excel workbook based on XML in C#
// Tags: XML-based freeze pane configuration for Aspose.Cells | C# parse FreezePane attributes from config.xml | set freeze rows and columns on first worksheet Aspose.Cells | programmatic Excel freeze panes using .NET API | load workbook and apply freeze pane settings via XML

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The example loads a config.xml file, extracts the FreezePane element's Row and Column attributes (defaulting to 0), opens input.xlsx with Aspose.Cells, applies those values to the first worksheet using the FreezePanes method, and saves the result as output.xlsx while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML configuration file
            string configPath = "config.xml";

            // Load the XML configuration if it exists
            int freezeRow = 0;
            int freezeColumn = 0;
            if (File.Exists(configPath))
            {
                XDocument configDoc = XDocument.Load(configPath);
                XElement freezeElement = configDoc.Root?.Element("FreezePane");

                // Read freeze row and column values (default to 0 if not specified)
                freezeRow = (int?)freezeElement?.Attribute("Row") ?? 0;
                freezeColumn = (int?)freezeElement?.Attribute("Column") ?? 0;
            }
            else
            {
                Console.WriteLine($"Configuration file '{configPath}' not found. Using default freeze values (0,0).");
            }

            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Ensure the input workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply freeze panes to the first worksheet using the 4‑parameter overload
            Worksheet sheet = workbook.Worksheets[0];
            sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
