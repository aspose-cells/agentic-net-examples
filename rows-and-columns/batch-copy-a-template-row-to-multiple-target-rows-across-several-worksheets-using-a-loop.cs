// Title: Copy a template row to multiple rows on several worksheets with Aspose.Cells for .NET (C#)
// Common Searches: how to duplicate a template row to many rows in different sheets using Aspose.Cells C# | batch copy rows across multiple worksheets with a loop in Aspose.Cells .NET | preserve cell formatting and row height when copying rows between worksheets in C# | create worksheets on the fly while copying rows with Aspose.Cells
// Tags: copy template row across worksheets Aspose.Cells | duplicate row values and formatting C# | batch copy rows with loop Aspose.Cells | preserve row height Aspose.Cells | auto‑create worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, extracts a template row, and copies its cell values, styles, and row height to multiple target rows defined for several worksheets. Missing worksheets are created automatically, and the updated workbook is saved.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWorkbook.xlsx";
                const string resultPath = "ResultWorkbook.xlsx";

                // Ensure the template file exists.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file \"{templatePath}\" not found.");
                    return;
                }

                // Load the workbook that contains the template row.
                Workbook workbook = new Workbook(templatePath);

                // Name of the worksheet that holds the template row.
                string templateSheetName = "TemplateSheet";
                // Zero‑based index of the template row to copy (e.g., third row).
                int templateRowIndex = 2;

                // Get the template worksheet; abort if missing.
                Worksheet templateSheet = workbook.Worksheets[templateSheetName];
                if (templateSheet == null)
                {
                    Console.WriteLine($"Worksheet \"{templateSheetName}\" not found in the template workbook.");
                    return;
                }

                // Determine the last column that contains data in the template row.
                int maxColumn = templateSheet.Cells.MaxColumn;

                // Define the target worksheets and the rows where the template should be copied.
                var copyPlan = new (string SheetName, int[] TargetRows)[]
                {
                    ("SheetA", new int[] { 5, 8, 12 }),
                    ("SheetB", new int[] { 3, 7 }),
                    ("SheetC", new int[] { 1, 4, 9, 15 })
                };

                // Loop through each target worksheet and each target row.
                foreach (var entry in copyPlan)
                {
                    // Get existing worksheet or create a new one if it does not exist.
                    Worksheet targetSheet = workbook.Worksheets[entry.SheetName] ?? workbook.Worksheets.Add(entry.SheetName);

                    foreach (int targetRowIndex in entry.TargetRows)
                    {
                        // Copy each cell from the template row to the target row.
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell srcCell = templateSheet.Cells[templateRowIndex, col];
                            Cell destCell = targetSheet.Cells[targetRowIndex, col];

                            // Copy value.
                            destCell.PutValue(srcCell.Value);

                            // Copy style (font, color, borders, etc.).
                            destCell.SetStyle(srcCell.GetStyle());
                        }

                        // Copy row height to preserve appearance.
                        targetSheet.Cells.Rows[targetRowIndex].Height = templateSheet.Cells.Rows[templateRowIndex].Height;
                    }
                }

                // Save the modified workbook.
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
