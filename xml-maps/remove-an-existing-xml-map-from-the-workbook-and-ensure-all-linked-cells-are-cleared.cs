// Title: Remove XML Maps and Clear Linked Cells in Excel with Aspose.Cells for .NET (C#)
// Description: Load a workbook, erase values of cells bound to XML maps, clear the XmlMapCollection, and save the file without any XML mappings using Aspose.Cells for C#.
// Keywords: Aspose.Cells XML map removal | C# clear cells linked to XML map | delete XML maps from Excel workbook | Aspose.Cells clear used range | remove XmlMapCollection programmatically | Excel XML map cleanup .NET
// Common Searches: how to delete xml map in Aspose.Cells C# | clear cells bound to xml map using Aspose.Cells | remove all xml maps from an Excel file programmatically | Aspose.Cells example for xml map removal | C# code to purge xml maps from workbook
// Developer Intent: Programmatically eliminate every XML map from a workbook and reset the contents of cells that were previously bound to those maps.
// Use Cases: Sanitizing a template before reuse by stripping XML bindings and emptying data cells. | Preparing Excel files for distribution to users who do not require XML data connections. | Batch processing multiple workbooks to remove XML maps and clear associated cell values.
// AI Prompts: Generate C# code with Aspose.Cells that removes all XML maps from a workbook and clears the values of every used cell. | Explain how to iterate through worksheets, obtain the used range, clear cell contents, and then call XmlMapCollection.Clear() in Aspose.Cells. | Provide a modification to the sample that only clears cells belonging to a specific XML map instead of the entire used range.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsXmlMapRemovalDemo
{
    // Load a workbook, erase values of cells bound to XML maps, clear the XmlMapCollection, and save the file without any XML mappings using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "InputWithXmlMap.xlsx";
                const string outputPath = "OutputWithoutXmlMap.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load an existing workbook that contains XML maps
                Workbook workbook = new Workbook(inputPath);

                // Store reference to the XmlMap collection
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // If there are any XML maps, remove them
                if (xmlMaps.Count > 0)
                {
                    // Iterate through each worksheet and clear cells that were linked to any XML map
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Get the used range of the worksheet
                        AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                        // If the worksheet is empty, skip clearing
                        if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                            continue;

                        int startRow = usedRange.FirstRow;
                        int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                        int startCol = usedRange.FirstColumn;
                        int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                        // Loop through each cell in the used range and clear its value
                        for (int row = startRow; row <= endRow; row++)
                        {
                            for (int col = startCol; col <= endCol; col++)
                            {
                                sheet.Cells[row, col].PutValue(string.Empty);
                            }
                        }
                    }

                    // Remove all XML maps from the workbook
                    xmlMaps.Clear();
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved without XML maps to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
