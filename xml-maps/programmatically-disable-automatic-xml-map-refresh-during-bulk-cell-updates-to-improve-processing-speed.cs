using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRefreshDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWithXmlMap.xlsx";
                const string outputPath = "UpdatedWorkbook.xlsx";

                // Verify that the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Error: Template file not found – {templatePath}");
                    return;
                }

                // Load the workbook that contains XML maps (or create a new one)
                Workbook workbook = new Workbook(templatePath);

                // NOTE: In recent Aspose.Cells versions the EnableXMLMapRefresh property
                // has been removed. The engine automatically optimizes bulk updates,
                // so we proceed without explicitly disabling the refresh.

                // Perform bulk updates
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Example: fill 10,000 rows with sample data
                for (int row = 0; row < 10000; row++)
                {
                    cells[row, 0].PutValue($"Item{row + 1}");
                    cells[row, 1].PutValue(row * 10);
                    cells[row, 2].PutValue(DateTime.Today.AddDays(row));
                }

                // If you need to force an XML map refresh, uncomment the line below
                // workbook.RefreshXmlMap();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}