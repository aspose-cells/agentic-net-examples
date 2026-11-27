// Title: Compare XML Maps in a Single Workbook to Detect Overlapping Table Ranges and Resolve Conflicts Using Aspose.Cells for .NET (C#)
// Common Searches: how to compare two XML maps in the same Excel workbook with Aspose.Cells C# | detect overlapping ListObject ranges from different XmlMaps using Aspose.Cells | resolve cell assignment conflicts between multiple XML maps in an Excel file C# | programmatically move an Excel table to avoid XmlMap overlap Aspose.Cells | C# code to check and fix overlapping XML map tables in a workbook
// Tags: xml map overlapping range detection Aspose.Cells | listobject conflict resolution with XmlMap C# | adjust table position to prevent XmlMap collision | compare multiple XmlMaps for cell assignment overlap | programmatic xml map range fix in Excel workbook

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace XmlMapOverlapResolver
{
    // The example loads an Excel workbook, gathers all ListObjects linked to XmlMaps, and compares each pair of tables that use different XmlMaps to find overlapping cell ranges. When an overlap is detected, it calculates a row offset, inserts rows if needed, and resizes the second table to move it below the first, updating its coordinates for further checks. The resolved workbook is then saved.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains two (or more) XML maps.
            // Ensure the file exists; otherwise an exception will be thrown.
            const string inputPath = "input.xlsx";
            const string outputPath = "output_resolved.xlsx";

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Workbook '{inputPath}' loaded successfully.");

                // Collect all ListObjects (tables) that are linked to an XmlMap.
                // A ListObject may have an XmlMap via its XmlMap property.
                List<ListObjectInfo> mappedTables = new List<ListObjectInfo>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // If the table is linked to an XmlMap, store its info.
                        if (table.XmlMap != null)
                        {
                            mappedTables.Add(new ListObjectInfo
                            {
                                Worksheet = sheet,
                                Table = table,
                                XmlMapName = table.XmlMap.Name,
                                StartRow = table.StartRow,
                                EndRow = table.EndRow,
                                StartColumn = table.StartColumn,
                                EndColumn = table.EndColumn
                            });

                            Console.WriteLine($"Found mapped table '{table.DisplayName}' on sheet '{sheet.Name}' using XmlMap '{table.XmlMap.Name}'.");
                        }
                    }
                }

                // No mapped tables? Nothing to do.
                if (mappedTables.Count < 2)
                {
                    Console.WriteLine("Less than two tables are linked to XML maps. No overlap check required.");
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved as '{outputPath}'.");
                    return;
                }

                // Detect overlapping assignments between tables that use different XmlMaps.
                // We'll compare each pair only once.
                bool anyOverlap = false;
                for (int i = 0; i < mappedTables.Count; i++)
                {
                    for (int j = i + 1; j < mappedTables.Count; j++)
                    {
                        ListObjectInfo a = mappedTables[i];
                        ListObjectInfo b = mappedTables[j];

                        // Skip if both tables use the same XmlMap – they are intentionally sharing the same mapping.
                        if (a.XmlMapName.Equals(b.XmlMapName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (AreasOverlap(a, b))
                        {
                            anyOverlap = true;
                            Console.WriteLine($"Overlap detected between table '{a.Table.DisplayName}' (Map: {a.XmlMapName}) " +
                                              $"and table '{b.Table.DisplayName}' (Map: {b.XmlMapName}).");

                            // Resolve conflict by moving the second table (b) downwards enough to avoid overlap.
                            // Compute the required offset.
                            int offsetRows = (a.EndRow - b.StartRow) + 1; // at least one row gap
                            int newStartRow = b.StartRow + offsetRows;
                            int newEndRow = b.EndRow + offsetRows;

                            // Ensure we don't exceed worksheet limits.
                            int maxRows = b.Worksheet.Cells.MaxDataRow + 1;
                            if (newEndRow > maxRows)
                            {
                                // If moving down would exceed current data range, expand the worksheet.
                                b.Worksheet.Cells.InsertRows(b.EndRow + 1, offsetRows);
                                Console.WriteLine($"Inserted {offsetRows} rows to accommodate moved table.");
                            }

                            // Perform the move.
                            b.Table.Resize(newStartRow, b.StartColumn, newEndRow - newStartRow + 1, b.EndColumn - b.StartColumn + 1, true);
                            Console.WriteLine($"Table '{b.Table.DisplayName}' moved to rows {newStartRow}-{newEndRow}.");

                            // Update stored coordinates for further comparisons.
                            b.StartRow = newStartRow;
                            b.EndRow = newEndRow;
                        }
                    }
                }

                if (!anyOverlap)
                {
                    Console.WriteLine("No overlapping cell assignments were found between different XML maps.");
                }

                // Save the resolved workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Resolved workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        // Helper method to determine if two rectangular cell areas overlap.
        private static bool AreasOverlap(ListObjectInfo a, ListObjectInfo b)
        {
            bool rowsOverlap = a.StartRow <= b.EndRow && a.EndRow >= b.StartRow;
            bool colsOverlap = a.StartColumn <= b.EndColumn && a.EndColumn >= b.StartColumn;
            return rowsOverlap && colsOverlap;
        }

        // Simple container for the information we need about each mapped ListObject.
        private class ListObjectInfo
        {
            public Worksheet Worksheet { get; set; }
            public ListObject Table { get; set; }
            public string XmlMapName { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartColumn { get; set; }
            public int EndColumn { get; set; }
        }
    }
}
