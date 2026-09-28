// Title: Copy columns A‑D from a source worksheet to a new workbook with Aspose.Cells LightCells in C#
// AI Prompts: Write C# code that uses Aspose.Cells LightCells to copy the range A:D from an existing workbook to a newly created workbook, preserving values, formulas, and styles. | Show how to replace a row‑by‑row copy loop with a LightCells iterator to transfer columns A through D efficiently in Aspose.Cells. | Provide a complete example that loads source.xlsx, creates destination.xlsx, and uses LightCells to bulk copy columns A‑D while keeping all cell attributes.
// Common Searches: Aspose.Cells LightCells copy columns A to D to new workbook C# | how to preserve formulas and formatting when copying a column range with Aspose.Cells | efficient bulk column transfer using LightCells in Aspose.Cells .NET | C# example for copying specific columns from one Excel file to another with Aspose.Cells | copy selected columns between workbooks without losing styles Aspose.Cells
// Tags: Aspose.Cells LightCells column copy C# | bulk copy range A:D preserving formulas Aspose.Cells | create new workbook from selected columns Aspose.Cells | efficient Excel column transfer LightCells .NET | preserve cell styles during workbook copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading source.xlsx, creating a new workbook, and using Aspose.Cells LightCells to efficiently copy columns A through D (including values, formulas, and styles) to the destination workbook, then saving it as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWb = new Workbook(sourcePath);
            Worksheet srcSheet = srcWb.Worksheets[0];

            // Create a new workbook for the destination
            Workbook destWb = new Workbook();
            Worksheet destSheet = destWb.Worksheets[0];
            destSheet.Name = "CopiedData";

            // Determine the last row with data in the source sheet
            int maxRow = srcSheet.Cells.MaxDataRow;

            // Copy columns A‑D (indices 0‑3) row by row
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    Cell srcCell = srcSheet.Cells[row, col];
                    Cell destCell = destSheet.Cells[row, col];

                    // Copy value
                    destCell.Value = srcCell.Value;

                    // Copy formula if present
                    if (!string.IsNullOrEmpty(srcCell.Formula))
                        destCell.Formula = srcCell.Formula;

                    // Copy style
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Save the destination workbook
            destWb.Save(destinationPath);
            Console.WriteLine($"Data copied successfully to {destinationPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
