// Title: Create a C# Aspose.Cells utility that maps and highlights circular reference chains in an Excel workbook
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, calls Workbook.CalculateFormula(), retrieves circular reference chains via reflection, and records each chain’s number, cell address, and order on a new worksheet. | Enhance the utility to apply a yellow background style to every cell that participates in a circular reference on its original sheet. | Add hyperlinks from the mapping worksheet back to each source cell so users can jump directly to the circular reference locations.
// Common Searches: aspnet detect circular reference chains in Excel using Aspose.Cells | C# Aspose.Cells visualize formula circular dependencies | highlight cells involved in circular references Aspose.Cells .NET | export circular reference list to new worksheet with Aspose.Cells | use reflection to access CircularReferenceInfo in Aspose.Cells C#
// Tags: Aspose.Cells circular reference visualization | C# extract circular reference chains | highlight circular reference cells Aspose.Cells | write reference chain worksheet Aspose.Cells | reflection access CircularReferenceInfo Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, forces formula calculation to capture circular reference data, uses reflection to obtain each circular reference chain, writes chain number, cell address, and order to a new worksheet, highlights the original cells with a yellow background, and saves the updated file.
class CircularReferenceVisualizer
{
    static void Main(string[] args)
    {
        // Paths for the source workbook and the result workbook.
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Force calculation to populate circular reference information.
            workbook.CalculateFormula();

            // Try to obtain circular reference information via reflection (API may vary by version).
            object circularInfo = null;
            var prop = workbook.GetType().GetProperty("CircularReferenceInfo");
            if (prop != null)
            {
                circularInfo = prop.GetValue(workbook);
            }

            // Create a new worksheet to visualize the circular reference chains.
            Worksheet mapSheet = workbook.Worksheets.Add("CircularReferenceMap");
            Cells mapCells = mapSheet.Cells;

            // Write header row.
            mapCells["A1"].PutValue("Chain #");
            mapCells["B1"].PutValue("Cell Address");
            mapCells["C1"].PutValue("Order in Chain");

            int currentRow = 1; // Zero‑based index; start after header.

            if (circularInfo != null)
            {
                // Access the collection of items (dynamic to avoid compile‑time dependency).
                var collectionProp = circularInfo.GetType().GetProperty("CircularReferenceInfoCollection");
                var collection = collectionProp?.GetValue(circularInfo) as System.Collections.IEnumerable;

                if (collection != null)
                {
                    foreach (var item in collection)
                    {
                        // Each item should have a 'Chain' property like "A1->B2->C3".
                        var chainProp = item.GetType().GetProperty("Chain");
                        string chain = chainProp?.GetValue(item) as string ?? string.Empty;

                        string[] chainCells = chain.Split(new[] { "->" }, StringSplitOptions.RemoveEmptyEntries);
                        int chainIndex = ++currentRow; // Unique index for this chain.

                        for (int i = 0; i < chainCells.Length; i++)
                        {
                            // Write mapping information to the visualization sheet.
                            mapCells[currentRow, 0].PutValue(chainIndex);               // Chain #
                            mapCells[currentRow, 1].PutValue(chainCells[i]);           // Cell address
                            mapCells[currentRow, 2].PutValue(i + 1);                   // Order in chain

                            // Determine sheet name and cell address.
                            string sheetName = workbook.Worksheets[0].Name; // Default to first sheet.
                            string address = chainCells[i];

                            if (address.Contains("!"))
                            {
                                var parts = address.Split('!');
                                sheetName = parts[0];
                                address = parts[1];
                            }

                            try
                            {
                                Worksheet sourceSheet = workbook.Worksheets[sheetName];
                                Cell sourceCell = sourceSheet.Cells[address];

                                // Apply a yellow background to indicate participation in a circular reference.
                                Style style = sourceCell.GetStyle();
                                style.ForegroundColor = Color.Yellow;
                                style.Pattern = BackgroundType.Solid;
                                sourceCell.SetStyle(style);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Failed to style cell '{address}' on sheet '{sheetName}': {ex.Message}");
                            }

                            currentRow++;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Circular reference collection is empty or unavailable.");
                }
            }
            else
            {
                Console.WriteLine("Circular reference information is not supported in this Aspose.Cells version.");
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
