// Title: C# – Highlight Cells Above Average in Column F Using Aspose.Cells Conditional Formatting
// Description: This example creates a new workbook, fills column F (rows 0‑9) with numeric values, adds an AboveAverage conditional format, sets a yellow background for cells that exceed the column average, and saves the result as an XLSX file.
// Keywords: Aspose.Cells C# conditional formatting | AboveAverage format Aspose.Cells | highlight cells above average column F | set cell background color Aspose.Cells | C# Excel conditional formatting example | Aspose.Cells conditional format range
// Common Searches: Aspose.Cells highlight values above average C# | Conditional formatting column F Aspose.Cells .NET | Apply AboveAverage rule with Aspose.Cells | C# code to color cells greater than average in Excel | How to use FormatConditionType.AboveAverage in Aspose.Cells
// Developer Intent: Generate a workbook and apply an AboveAverage conditional format to column F, coloring qualifying cells yellow.
// Use Cases: Sales dashboards where figures above the column average are instantly visible. | Academic grade sheets that flag scores higher than the class mean. | Expense reports that draw attention to outlier amounts in a specific column.
// AI Prompts: Write C# code with Aspose.Cells to apply an AboveAverage conditional format to column F and use a yellow fill. | Show how to change the formatted range to rows 0‑20 in column F and switch the fill color to light green. | Explain how to compute the average of column F manually and create a custom rule that highlights cells greater than that value.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingDemo
{
    // This example creates a new workbook, fills column F (rows 0‑9) with numeric values, adds an AboveAverage conditional format, sets a yellow background for cells that exceed the column average, and saves the result as an XLSX file.
    public class HighlightAboveAverageInColumnF
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data in column F (index 5)
                for (int row = 0; row < 10; row++)
                {
                    // Example values: 10, 20, ..., 100
                    worksheet.Cells[row, 5].PutValue((row + 1) * 10);
                }

                // Add a conditional formatting collection to the worksheet
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection fcc = worksheet.ConditionalFormattings[cfIndex];

                // Define the range: column F rows 0 through 9
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    EndRow = 9,
                    StartColumn = 5,
                    EndColumn = 5
                };
                fcc.AddArea(area);

                // Add an AboveAverage condition to the collection
                int conditionIndex = fcc.AddCondition(FormatConditionType.AboveAverage);
                FormatCondition fc = fcc[conditionIndex];

                // Configure the condition to highlight values above the average
                fc.AboveAverage.IsAboveAverage = true;          // true = above average
                fc.Style.BackgroundColor = Color.Yellow;        // highlight color

                // Save the workbook
                string outputPath = "HighlightAboveAverageInColumnF.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HighlightAboveAverageInColumnF.Run();
        }
    }
}
