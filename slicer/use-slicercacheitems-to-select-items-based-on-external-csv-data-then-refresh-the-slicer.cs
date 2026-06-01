using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCsvDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string workbookPath = "input.xlsx";
                const string csvPath = "filter.csv";
                const string outputPath = "output.xlsx";

                // Verify input files exist
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"CSV file not found: {csvPath}");
                    return;
                }

                // Load workbook containing a pivot table and slicer
                Workbook workbook = new Workbook(workbookPath);
                // Refresh all data connections (pivot tables, slicers, etc.)
                workbook.Worksheets.RefreshAll();

                // Read filter values from CSV (one value per line)
                string[] csvValues = File.ReadAllLines(csvPath)
                                         .Select(line => line.Trim())
                                         .Where(line => !string.IsNullOrEmpty(line))
                                         .ToArray();

                // Locate the first slicer in the workbook
                Slicer slicer = null;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Slicers.Count > 0)
                    {
                        slicer = ws.Slicers[0];
                        break;
                    }
                }

                if (slicer == null)
                {
                    Console.WriteLine("No slicer found in the workbook.");
                    return;
                }

                // Access slicer cache items
                SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;

                // Deselect all items
                foreach (SlicerCacheItem item in cacheItems)
                {
                    item.Selected = false;
                }

                // Select items matching CSV values
                foreach (string value in csvValues)
                {
                    SlicerCacheItem match = cacheItems.FirstOrDefault(i =>
                        i.Value.Equals(value, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        match.Selected = true;
                    }
                }

                // Refresh slicer (updates underlying pivot table)
                slicer.Refresh();

                // Save modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}