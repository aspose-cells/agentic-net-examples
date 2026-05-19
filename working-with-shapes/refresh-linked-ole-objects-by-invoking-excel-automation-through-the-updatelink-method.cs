using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleRefreshDemo
{
    class Program
    {
        static void Main()
        {
            const string inputFile = "InputWithOleLinks.xlsx";
            const string outputFile = "OutputWithRefreshedOleLinks.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook that contains linked OLE objects
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all OLE objects in the current worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // Process only linked OLE objects
                        if (ole.IsLink)
                        {
                            // Ensure the OLE object is set to auto‑update when the source changes
                            ole.AutoUpdate = true;

                            // Display the source file path for debugging
                            Console.WriteLine($"Refreshing OLE link: {ole.ObjectSourceFullName}");
                        }
                    }
                }

                // Refresh all linked data sources (including linked OLE objects)
                workbook.UpdateLinkedDataSource(new Workbook[0]);

                // Save the workbook after refreshing the OLE links
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                // Catch any runtime errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}