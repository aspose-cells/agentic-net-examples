// Title: Apply freeze pane rows and columns from an XML configuration file to an Aspose.Cells workbook in C#
// AI Prompts: Read the Row and Column elements from a config.xml file and use them to call Worksheet.FreezePanes on the target worksheet. | Extend the sample to parse multiple <Worksheet> nodes in the XML, each containing its own Row and Column values, and apply FreezePanes to the corresponding sheets. | Add validation that the Row and Column numbers are within the worksheet's total rows and columns before invoking FreezePanes, and log a warning if they are out of range.
// Common Searches: c# aspose.cells set freeze panes using values from an external xml file | how to read freeze row and column from config.xml and apply to first worksheet in Aspose.Cells | aspnet read integer settings from xml and use FreezePanes method in Aspose.Cells | validate freeze pane coordinates against worksheet dimensions before calling FreezePanes in C#
// Tags: xml driven freeze pane settings Aspose.Cells | apply FreezePanes from configuration values C# | multiple worksheet freeze configuration xml | validate freeze pane coordinates worksheet size | error handling missing config file Aspose.Cells

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The program loads Row and Column values from a config.xml file, opens an existing Excel workbook with Aspose.Cells, applies FreezePanes to the first worksheet using those values, and saves the modified file, including checks for file existence and basic error handling.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Define file paths (adjust as needed)
            string configPath = "config.xml";               // XML with freeze settings
            string inputWorkbookPath = "input.xlsx";        // Existing workbook to modify
            string outputWorkbookPath = "output.xlsx";      // Destination workbook

            // Verify that the configuration file exists
            if (!File.Exists(configPath))
                throw new FileNotFoundException($"Configuration file not found: {configPath}");

            // Verify that the input workbook exists
            if (!File.Exists(inputWorkbookPath))
                throw new FileNotFoundException($"Input workbook not found: {inputWorkbookPath}");

            // Load the XML configuration file
            XDocument configDoc = XDocument.Load(configPath);

            // Retrieve freeze row and column values; default to 0 if elements are missing
            int freezeRows = (int?)configDoc.Root?.Element("Row") ?? 0;
            int freezeColumns = (int?)configDoc.Root?.Element("Column") ?? 0;

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(inputWorkbookPath);

            // Access the first worksheet (you can change the index or use a name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Apply freeze panes based on the configuration values
            // FreezePanes(totalRows, totalColumns, row, column)
            // The last two parameters specify the cell that will be the top‑left cell of the scrollable area.
            if (freezeRows > 0 || freezeColumns > 0)
            {
                worksheet.FreezePanes(freezeRows, freezeColumns, freezeRows, freezeColumns);
            }

            // Save the modified workbook
            workbook.Save(outputWorkbookPath);
            Console.WriteLine($"Workbook saved successfully to '{outputWorkbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
