// Title: Compare worksheet column widths, row heights, and merged cell ranges before and after saving to XML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, exports it to XML using Aspose.Cells, reloads the XML into a new Workbook object, and confirms that the layout—including column dimensions, row dimensions, and merged ranges—matches the original. | Create a C# helper method that receives two Aspose.Cells Workbook instances and returns a boolean indicating whether their layout properties are identical within a 0.001 tolerance.
// Common Searches: asp.net verify worksheet layout after saving to XML with Aspose.Cells | C# ensure column size and row size remain unchanged after Excel XML conversion | check if merged cell areas are retained when importing XML using Aspose.Cells | how to compare two workbooks for geometry consistency in .NET | tolerance settings for geometry comparison in Aspose.Cells C#
// Tags: Aspose.Cells worksheet geometry comparison | XML export column dimension verification | row dimension tolerance check Aspose.Cells | merged area validation C# | Excel layout integrity after XML import

using System;
using System.IO;
using Aspose.Cells;

// The program loads or creates an Excel workbook, saves it as XML with Aspose.Cells, reloads the XML into a new workbook, and compares column widths, row heights, and merged cell ranges between the original and imported workbooks using a small floating‑point tolerance.
class GeometryComparison
{
    static void Main(string[] args)
    {
        try
        {
            string originalPath = "original.xlsx";
            // Ensure the original file exists; create a simple workbook if it does not.
            if (!File.Exists(originalPath))
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample");
                wb.Save(originalPath);
            }

            // Load the original workbook.
            Workbook originalWorkbook = new Workbook(originalPath);

            // Export the workbook geometry to an intermediate XML file.
            string tempPath = "temp.xml";
            originalWorkbook.Save(tempPath, SaveFormat.Xml);

            // Load a new workbook from the exported XML.
            Workbook importedWorkbook = new Workbook(tempPath);

            // Compare geometry of the two workbooks.
            bool isConsistent = CompareGeometry(originalWorkbook, importedWorkbook);

            Console.WriteLine("Geometry consistent: " + isConsistent);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // Compares column widths, row heights and merged cells for all worksheets.
    private static bool CompareGeometry(Workbook wb1, Workbook wb2)
    {
        // Ensure the same number of worksheets.
        if (wb1.Worksheets.Count != wb2.Worksheets.Count)
            return false;

        const double tolerance = 0.001; // tolerance for floating‑point comparisons

        for (int sheetIndex = 0; sheetIndex < wb1.Worksheets.Count; sheetIndex++)
        {
            Worksheet ws1 = wb1.Worksheets[sheetIndex];
            Worksheet ws2 = wb2.Worksheets[sheetIndex];

            // ----- Compare column widths -----
            int maxColumn = Math.Max(ws1.Cells.MaxColumn, ws2.Cells.MaxColumn);
            for (int col = 0; col <= maxColumn; col++)
            {
                double width1 = ws1.Cells.GetColumnWidth(col);
                double width2 = ws2.Cells.GetColumnWidth(col);
                if (Math.Abs(width1 - width2) > tolerance)
                    return false;
            }

            // ----- Compare row heights -----
            int maxRow = Math.Max(ws1.Cells.MaxRow, ws2.Cells.MaxRow);
            for (int row = 0; row <= maxRow; row++)
            {
                double height1 = ws1.Cells.GetRowHeight(row);
                double height2 = ws2.Cells.GetRowHeight(row);
                if (Math.Abs(height1 - height2) > tolerance)
                    return false;
            }

            // ----- Compare merged cells -----
            CellArea[] merged1 = ws1.Cells.GetMergedAreas();
            CellArea[] merged2 = ws2.Cells.GetMergedAreas();

            if (merged1.Length != merged2.Length)
                return false;

            for (int i = 0; i < merged1.Length; i++)
            {
                string range1 = CellsHelper.CellIndexToName(merged1[i].StartRow, merged1[i].StartColumn) + ":" +
                                CellsHelper.CellIndexToName(merged1[i].EndRow, merged1[i].EndColumn);
                string range2 = CellsHelper.CellIndexToName(merged2[i].StartRow, merged2[i].StartColumn) + ":" +
                                CellsHelper.CellIndexToName(merged2[i].EndRow, merged2[i].EndColumn);

                if (!string.Equals(range1, range2, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
        }

        // All checks passed.
        return true;
    }
}
