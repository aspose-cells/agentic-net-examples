using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

namespace AsposeCellsFreezeFromXml
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the XML configuration file
                string configPath = "freezeConfig.xml";

                // Verify that the configuration file exists
                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"Configuration file '{configPath}' not found.");
                    return;
                }

                // Load the XML configuration
                XDocument configDoc = XDocument.Load(configPath);
                XElement root = configDoc.Root;

                // Extract freeze pane parameters (default to 0 if missing)
                int row = (int?)root.Element("Row") ?? 0;
                int column = (int?)root.Element("Column") ?? 0;
                int freezedRows = (int?)root.Element("FreezedRows") ?? 0;
                int freezedColumns = (int?)root.Element("FreezedColumns") ?? 0;

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Apply freeze panes if configuration is valid
                if ((row != 0 || column != 0) && (freezedRows != 0 || freezedColumns != 0))
                {
                    sheet.FreezePanes(row, column, freezedRows, freezedColumns);
                }
                else
                {
                    Console.WriteLine("Invalid freeze pane configuration; skipping FreezePanes call.");
                }

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved to '{outputPath}' with freeze panes applied.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}