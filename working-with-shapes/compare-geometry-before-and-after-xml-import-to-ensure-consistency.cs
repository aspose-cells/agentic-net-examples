// Title: Check worksheet column widths, row heights, and merged cells after Excel 2003 XML round‑trip with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells compare column widths before and after saving to Excel 2003 XML | how to validate row height consistency after loading XML workbook with Aspose.Cells | verify merged cell ranges remain unchanged after XML round‑trip using Aspose.Cells .NET | detect geometry differences in worksheets after Excel 2003 XML import with Aspose.Cells
// Tags: column width comparison Aspose.Cells XML | row height verification Aspose.Cells C# | merged cells consistency Aspose.Cells XML | worksheet geometry validation Aspose.Cells | Excel 2003 XML roundtrip integrity Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGeometryComparison
{
    // The example creates a workbook, sets explicit column widths, row heights, and merges a cell range, then saves it to the Excel 2003 XML format in memory. It reloads the XML into a new Workbook and programmatically compares the original and imported workbooks for worksheet count, column widths, row heights, and merged cell areas, outputting whether the geometry remains consistent.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a sample workbook with some geometry
                // -------------------------------------------------
                Workbook originalWb = new Workbook();
                Worksheet sheet = originalWb.Worksheets[0];
                sheet.Name = "Sample";

                // Set column widths
                sheet.Cells.SetColumnWidth(0, 20); // A
                sheet.Cells.SetColumnWidth(1, 30); // B

                // Set row heights
                sheet.Cells.SetRowHeight(0, 25); // 1
                sheet.Cells.SetRowHeight(1, 35); // 2

                // Add some data
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue(123);
                sheet.Cells["B2"].PutValue(456);

                // Merge cells (geometry element)
                sheet.Cells.Merge(2, 0, 2, 2); // Merge C3:D4 (rows 3-4, cols C-D)

                // -------------------------------------------------
                // 2. Save the workbook to XML (Excel 2003 XML format) in memory
                // -------------------------------------------------
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    originalWb.Save(xmlStream, SaveFormat.Xml);
                    xmlStream.Position = 0; // Reset stream position for reading

                    // -------------------------------------------------
                    // 3. Load a new workbook from the XML stream
                    // -------------------------------------------------
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
                    Workbook importedWb = new Workbook(xmlStream, loadOptions);

                    // -------------------------------------------------
                    // 4. Compare geometry between original and imported workbooks
                    // -------------------------------------------------
                    bool geometryMatches = CompareGeometry(originalWb, importedWb);
                    Console.WriteLine($"Geometry consistency check: {(geometryMatches ? "PASS" : "FAIL")}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static bool CompareGeometry(Workbook wb1, Workbook wb2)
        {
            bool isConsistent = true;

            // Compare worksheet count
            if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            {
                Console.WriteLine($"Worksheet count mismatch: {wb1.Worksheets.Count} vs {wb2.Worksheets.Count}");
                isConsistent = false;
            }

            // Iterate through worksheets
            int sheetCount = Math.Min(wb1.Worksheets.Count, wb2.Worksheets.Count);
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet ws1 = wb1.Worksheets[i];
                Worksheet ws2 = wb2.Worksheets[i];

                // Compare column widths (first 256 columns)
                for (int col = 0; col < 256; col++)
                {
                    double w1 = ws1.Cells.GetColumnWidth(col);
                    double w2 = ws2.Cells.GetColumnWidth(col);
                    if (Math.Abs(w1 - w2) > 0.001) // tolerance for floating point
                    {
                        Console.WriteLine($"Column width mismatch in sheet '{ws1.Name}', column {col}: {w1} vs {w2}");
                        isConsistent = false;
                    }
                }

                // Compare row heights (first 65535 rows)
                for (int row = 0; row < 65535; row++)
                {
                    double h1 = ws1.Cells.GetRowHeight(row);
                    double h2 = ws2.Cells.GetRowHeight(row);
                    if (Math.Abs(h1 - h2) > 0.001)
                    {
                        Console.WriteLine($"Row height mismatch in sheet '{ws1.Name}', row {row}: {h1} vs {h2}");
                        isConsistent = false;
                    }
                }

                // Compare merged cells using the recommended API
                CellArea[] mergedAreas1 = ws1.Cells.GetMergedAreas();
                CellArea[] mergedAreas2 = ws2.Cells.GetMergedAreas();

                if (mergedAreas1.Length != mergedAreas2.Length)
                {
                    Console.WriteLine($"Merged cells count mismatch in sheet '{ws1.Name}': {mergedAreas1.Length} vs {mergedAreas2.Length}");
                    isConsistent = false;
                }
                else
                {
                    for (int m = 0; m < mergedAreas1.Length; m++)
                    {
                        CellArea area1 = mergedAreas1[m];
                        CellArea area2 = mergedAreas2[m];
                        if (area1.StartRow != area2.StartRow ||
                            area1.StartColumn != area2.StartColumn ||
                            area1.EndRow != area2.EndRow ||
                            area1.EndColumn != area2.EndColumn)
                        {
                            Console.WriteLine($"Merged cell area mismatch in sheet '{ws1.Name}' at index {m}: " +
                                              $"[{area1.StartRow},{area1.StartColumn}] - [{area1.EndRow},{area1.EndColumn}] vs " +
                                              $"[{area2.StartRow},{area2.StartColumn}] - [{area2.EndRow},{area2.EndColumn}]");
                            isConsistent = false;
                        }
                    }
                }
            }

            return isConsistent;
        }
    }
}
