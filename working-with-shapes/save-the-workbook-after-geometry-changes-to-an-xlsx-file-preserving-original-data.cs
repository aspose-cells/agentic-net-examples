// Title: Save a modified workbook to a new XLSX file after adjusting column width and row height while preserving original data using Aspose.Cells for C#
// Common Searches: how to change column width and row height then save workbook as new file using Aspose.Cells C# | preserve existing worksheet data when saving geometry changes with Aspose.Cells | Aspose.Cells C# save workbook after column width adjustment without losing data | C# example to modify row height and export to new XLSX using Aspose.Cells | load Excel file, adjust geometry, and save new file with Aspose.Cells in .NET
// Tags: adjust column width Aspose.Cells | set row height Aspose.Cells | save workbook as XLSX Aspose.Cells | preserve original data Aspose.Cells | Excel geometry modification C#

using System;
using Aspose.Cells;

namespace AsposeCellsGeometryExample
{
    // The example loads an existing XLSX workbook, changes column A width to 20 characters and the first row height to 30 points, then saves the workbook to a new XLSX file, ensuring all original worksheet data remains intact.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (preserves all original data)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet (you can modify as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Example geometry changes:
            // Set column A width to 20 characters
            sheet.Cells.SetColumnWidth(0, 20);

            // Set row 1 height to 30 points
            sheet.Cells.SetRowHeight(0, 30);

            // You can add more geometry modifications here (e.g., merge cells, set column/row visibility, etc.)

            // Save the modified workbook to a new XLSX file while preserving original data
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
