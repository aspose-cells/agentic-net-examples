// Title: Apply red fill conditional formatting to cells with values > 100 after importing XML map data using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a workbook, imports an XML file via an Aspose.Cells XML map, defines a target range, and adds a conditional formatting rule that colors cells red when the value exceeds 100. | Show how to use dynamic objects in C# with Aspose.Cells to attach a cell‑value condition to a specific range after loading XML map data.
// Common Searches: asp.net c# how to import XML map with Aspose.Cells and set conditional formatting for values above 100 | conditional formatting based on XML map data in Aspose.Cells C# example | apply red background to cells greater than 100 after XML map import using Aspose.Cells | using dynamic objects to add conditional formatting in Aspose.Cells C# | create range and add cell value condition from imported XML data Aspose.Cells
// Tags: xml map import Aspose.Cells C# | conditional formatting cell value Aspose.Cells | apply red fill style Aspose.Cells | dynamic object Aspose.Cells API | range creation Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, imports data from a local XML file through an Aspose.Cells XML map named "MyMap", defines the range B2:B10, adds a conditional formatting rule that highlights cells with values greater than 100 in red, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Path to the XML file that contains the data
            string xmlPath = "data.xml";

            // Ensure the XML file exists before attempting to import
            if (File.Exists(xmlPath))
            {
                try
                {
                    // Use dynamic to call XML map APIs that may not be present in older versions
                    dynamic wbDynamic = workbook;
                    wbDynamic.XmlMaps.Add("MyMap", "Root", string.Empty);
                    wbDynamic.ImportXmlMap(xmlPath, "MyMap");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"XML import failed: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: XML file not found at path '{xmlPath}'. Skipping XML import.");
            }

            // Define the range where the XML data (or sample data) is located
            int firstRow = 1;      // Row index 1 => Excel row 2
            int firstColumn = 1;   // Column index 1 => Excel column B
            int totalRows = 9;     // Rows 2 through 10
            int totalColumns = 1;  // Single column

            // Create the range object (fully qualified to avoid ambiguity with System.Range)
            Aspose.Cells.Range range = sheet.Cells.CreateRange(firstRow, firstColumn, totalRows, totalColumns);

            // Add a conditional formatting rule to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();

            // Use dynamic to avoid compile‑time dependency on ConditionalFormatting type
            dynamic cf = sheet.ConditionalFormattings[cfIndex];
            cf.AddArea(range); // Apply to the defined range

            // Condition: cell value greater than 100
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "100",
                string.Empty);

            // Define the style for the condition
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Red;
            style.Pattern = BackgroundType.Solid;
            cf[conditionIndex].Style = style;

            // Save the workbook with the applied conditional formatting
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
