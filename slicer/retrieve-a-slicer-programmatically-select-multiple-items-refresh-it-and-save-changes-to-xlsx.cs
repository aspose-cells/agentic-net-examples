// Title: Programmatically select multiple slicer items, refresh the slicer, and save the updated XLSX workbook using Aspose.Cells for .NET
// Common Searches: C# code to select several items in an Excel slicer with Aspose.Cells | how to refresh a slicer after changing its selections using Aspose.Cells .NET | saving an XLSX file after updating slicer selections in C# | example of clearing and adding slicer items dynamically with Aspose.Cells | using dynamic to manipulate slicer objects in Aspose.Cells for .NET
// Tags: Aspose.Cells slicer selection manipulation | Aspose.Cells slicer refresh method | Aspose.Cells workbook save after slicer update | dynamic slicer access Aspose.Cells | C# programmatic slicer item addition

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX workbook, checks for slicers on the first worksheet, clears any prior selections on the first slicer, adds three specific items ("Item A", "Item B", "Item C"), refreshes the slicer to apply the changes, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure worksheet contains slicers
            if (worksheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the worksheet.");
                return;
            }

            // Retrieve first slicer using dynamic to avoid compile‑time dependency on Slicer type
            dynamic slicer = worksheet.Slicers[0];

            try
            {
                // Clear previous selections
                slicer.SelectedItems.Clear();

                // Add desired items
                slicer.SelectedItems.Add("Item A");
                slicer.SelectedItems.Add("Item B");
                slicer.SelectedItems.Add("Item C");

                // Apply changes
                slicer.Refresh();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Slicer operation error: {ex.Message}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
