// Title: Refresh a linked SmartArt shape after changing its source cell with Aspose.Cells for .NET
// Description: Loads or creates a workbook, updates cell B2, calls worksheet.Shapes.UpdateSelectedValue, and saves the file using OoxmlSaveOptions.UpdateSmartArt so the SmartArt diagram reflects the new value.
// Keywords: Aspose.Cells | C# | SmartArt refresh | UpdateSelectedValue | UpdateSmartArt | linked shape refresh | Shape.RefreshLinkedShape | worksheet.Shapes | Excel automation | programmatic SmartArt update
// Common Searches: how to refresh SmartArt after cell change Aspose.Cells | Aspose.Cells .NET update linked SmartArt diagram | worksheet.Shapes.UpdateSelectedValue example | Enable UpdateSmartArt when saving workbook | refresh linked shape programmatically Excel C#
// Developer Intent: Refresh a SmartArt graphic that is linked to worksheet cells after the cell values have been modified.
// Use Cases: Keep KPI dashboards up‑to‑date by refreshing SmartArt diagrams after data imports. | Automate monthly reports that modify worksheet values and need refreshed SmartArt before distribution. | Batch‑process workbooks to recalculate all linked SmartArt graphics after a data migration.
// AI Prompts: Show me C# code that uses Shape.RefreshLinkedShape or worksheet.Shapes.UpdateSelectedValue to refresh a SmartArt shape after updating its source cell with Aspose.Cells. | Provide an example that updates multiple linked SmartArt shapes and saves the workbook with UpdateSmartArt enabled. | Explain the difference between Shape.RefreshLinkedShape and worksheet.Shapes.UpdateSelectedValue in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads or creates a workbook, updates cell B2, calls worksheet.Shapes.UpdateSelectedValue, and saves the file using OoxmlSaveOptions.UpdateSmartArt so the SmartArt diagram reflects the new value.
    class RefreshSmartArtExample
    {
        static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input workbook exists; create a minimal workbook if it does not.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a default worksheet and a placeholder SmartArt shape if needed.
                Worksheet ws = workbook.Worksheets[0];
                ws.Name = "Sheet1";
                ws.Cells["B2"].PutValue("Initial Value");
                // Note: Adding actual SmartArt programmatically is beyond this example.
                workbook.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook at '{inputPath}'.");
            }

            // Update the target cell that the SmartArt is linked to
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["B2"].PutValue("New Value");

            // Refresh the linked shape values (including SmartArt) after the cell change
            worksheet.Shapes.UpdateSelectedValue();

            // Save the workbook with SmartArt update enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateSmartArt = true
            };
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved with updated SmartArt to '{outputPath}'.");
        }
    }
}
