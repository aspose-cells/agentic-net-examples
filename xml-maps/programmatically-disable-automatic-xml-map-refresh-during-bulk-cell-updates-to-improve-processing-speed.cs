using System;
using System.IO;
using Aspose.Cells;

class DisableXmlMapRefreshDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook wb = new Workbook(inputPath);

            // NOTE: In some Aspose.Cells versions the property
            // WorkbookSettings.EnableXMLMapRefresh is not available.
            // If it exists, you can uncomment the following lines:
            // wb.Settings.EnableXMLMapRefresh = false;

            // Perform bulk cell updates
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue(row * col);
                }
            }

            // Re‑enable XML map refresh after updates (if the property is supported)
            // wb.Settings.EnableXMLMapRefresh = true;

            // Save the modified workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}