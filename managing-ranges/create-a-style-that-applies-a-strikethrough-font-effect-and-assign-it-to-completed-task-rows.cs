// Title: Aspose.Cells .NET – Apply Strikethrough Font to Completed Task Rows
// Description: This example creates a workbook, defines a style with Font.IsStrikeout enabled, uses a StyleFlag to affect only the strike‑out property, scans a boolean "Completed" column, and applies the style to every entire row marked true before saving the file.
// Keywords: Aspose.Cells strikethrough style | C# apply font strikeout | StyleFlag font strikeout | conditional row formatting Aspose.Cells | Excel strikeout rows .NET
// Common Searches: Aspose.Cells apply strikethrough to rows | C# strikeout font based on boolean column | StyleFlag only font strikeout Aspose | how to format completed tasks in Excel with Aspose.Cells | conditional row style Aspose.Cells .NET
// Developer Intent: Create a strike‑out font style and automatically apply it to rows where the Completed column is true.
// Use Cases: Visually cross out finished tasks in a project tracker. | Produce printable task lists where completed items are clearly marked. | Highlight rows that satisfy a boolean condition while preserving other cell formats.
// AI Prompts: Write C# code using Aspose.Cells to add a strikeout font style to rows whose column B contains true, limiting changes to the font with StyleFlag. | Provide a reusable method that receives a worksheet, a boolean column index, and applies strikethrough formatting to all matching rows. | Explain how to replace the manual loop with Aspose.Cells conditional formatting to automatically strike through rows based on a boolean cell.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, defines a style with Font.IsStrikeout enabled, uses a StyleFlag to affect only the strike‑out property, scans a boolean "Completed" column, and applies the style to every entire row marked true before saving the file.
    public class StrikethroughCompletedRows
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data: first column contains task description,
                // second column indicates completion status (true = completed)
                cells["A1"].PutValue("Task");
                cells["B1"].PutValue("Completed");
                cells["A2"].PutValue("Design UI");
                cells["B2"].PutValue(true);
                cells["A3"].PutValue("Implement backend");
                cells["B3"].PutValue(false);
                cells["A4"].PutValue("Write tests");
                cells["B4"].PutValue(true);
                cells["A5"].PutValue("Deploy");
                cells["B5"].PutValue(false);

                // Create a style that applies a single strikethrough
                Style strikeStyle = workbook.CreateStyle();
                strikeStyle.Font.IsStrikeout = true; // enable strikeout on the font

                // Create a style flag indicating that only the FontStrike property should be applied
                StyleFlag flag = new StyleFlag();
                flag.FontStrike = true;

                // Apply the strikethrough style to rows where the task is completed
                // Rows are zero‑based; data starts at row index 1 (after header)
                for (int row = 1; row <= 4; row++)
                {
                    // Check the "Completed" column (index 1) for a true value
                    object cellValue = cells[row, 1].Value;
                    if (cellValue is bool isCompleted && isCompleted)
                    {
                        // Apply the style to the entire row
                        worksheet.Cells.ApplyRowStyle(row, strikeStyle, flag);
                    }
                }

                // Save the workbook
                string outputPath = "CompletedTasksStrikethrough.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StrikethroughCompletedRows.Run();
        }
    }
}
