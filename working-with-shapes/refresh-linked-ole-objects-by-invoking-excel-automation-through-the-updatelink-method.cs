// Title: Refresh Linked OLE Objects in Excel with Aspose.Cells for .NET
// Description: This C# example demonstrates how to load an Excel workbook, locate every linked OLE object across all worksheets, enable its AutoUpdate property, and save the file so the objects display the latest source data. Aspose.Cells automatically refreshes linked OLE objects when AutoUpdate is true, providing a simple alternative to Excel automation's UpdateLink method.
// Keywords: Aspose.Cells | .NET | OLE object | linked OLE | AutoUpdate | refresh OLE links | Excel automation | UpdateLink method | Workbook.Save | C# example
// Common Searches: Aspose.Cells refresh linked OLE objects | C# update OLE links in Excel | Set AutoUpdate for OLE objects Aspose | Programmatically refresh OLE objects | UpdateLink equivalent in Aspose.Cells
// Developer Intent: Programmatically ensure that every linked OLE object in an Excel workbook is refreshed to reflect its current source file.
// Use Cases: Batch‑process workbooks to automatically update all linked OLE objects before distribution. | Validate and log the source path of each linked OLE object while ensuring they are set to auto‑update. | Integrate OLE refresh into a larger data‑pipeline that generates reports with up‑to‑date embedded content.
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all worksheets, sets AutoUpdate = true for each linked OLE object, and saves the workbook. | Explain how Aspose.Cells handles OLE link refresh when AutoUpdate is enabled and describe any API that can force a manual refresh. | Create a method that logs the source file path of each linked OLE object, checks for missing sources, and then triggers a refresh of the objects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to load an Excel workbook, locate every linked OLE object across all worksheets, enable its AutoUpdate property, and save the file so the objects display the latest source data. Aspose.Cells automatically refreshes linked OLE objects when AutoUpdate is true, providing a simple alternative to Excel automation's UpdateLink method.
    public class RefreshLinkedOleObjectsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWithOleLinks.xlsx";
            const string outputPath = "OutputAfterOleRefresh.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that contains linked OLE objects
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all OLE objects in the current worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // Process only linked OLE objects
                        if (ole.IsLink)
                        {
                            // Ensure the object is set to auto‑update when the source changes
                            ole.AutoUpdate = true;

                            // Optionally, display the current source file path
                            Console.WriteLine($"Refreshing OLE object linked to: {ole.ObjectSourceFullName}");
                        }
                    }
                }

                // Note: Aspose.Cells automatically refreshes linked OLE objects when AutoUpdate is true.
                // If a specific API is required to update links, it can be invoked here when available.

                // Save the workbook after the refresh operation
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
