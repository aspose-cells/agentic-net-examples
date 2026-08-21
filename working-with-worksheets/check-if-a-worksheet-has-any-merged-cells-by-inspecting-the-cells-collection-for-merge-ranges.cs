// Title: Check for Merged Cells in a Worksheet with Aspose.Cells for .NET
// Description: Shows how to call Cells.GetMergedAreas() to find any merged ranges in a worksheet, return a true/false flag, and enumerate each merged area's coordinates. Includes an optional merge example and workbook save step.
// Keywords: Aspose.Cells | C# | GetMergedAreas | merged cells detection | worksheet merged ranges | Excel merge check | Aspose.Cells API | C# Excel processing
// Common Searches: Aspose.Cells find merged cells C# | GetMergedAreas example .NET | how to detect merged ranges in Excel using Aspose | check if worksheet has merged cells Aspose.Cells | list merged areas Aspose.Cells C#
// Developer Intent: Determine whether a worksheet contains any merged cells and obtain their locations.
// Use Cases: Validate an incoming workbook before data extraction to avoid parsing errors caused by merged cells. | Generate a report of merged ranges for auditing or documentation purposes. | Trigger conditional formatting or layout adjustments only when merged cells are present.
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, lists all merged cell ranges, and returns them as a collection. | Provide a C# snippet that unmerges every merged area in a worksheet after detecting them with Aspose.Cells. | Explain how to check for merged cells efficiently without loading the entire workbook into memory using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsMergedCheck
{
    // Shows how to call Cells.GetMergedAreas() to find any merged ranges in a worksheet, return a true/false flag, and enumerate each merged area's coordinates. Includes an optional merge example and workbook save step.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example merge (can be removed if checking an existing file)
            cells.Merge(0, 0, 2, 2); // Merges A1:B2

            // Retrieve all merged areas
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Determine if any merged cells exist
            bool hasMerged = mergedAreas != null && mergedAreas.Length > 0;

            Console.WriteLine("Worksheet has merged cells: " + hasMerged);
            if (hasMerged)
            {
                foreach (CellArea area in mergedAreas)
                {
                    Console.WriteLine($"Merged area: Row[{area.StartRow}-{area.EndRow}], Column[{area.StartColumn}-{area.EndColumn}]");
                }
            }

            // Save the workbook (optional)
            workbook.Save("MergedCheckResult.xlsx");
        }
    }
}
