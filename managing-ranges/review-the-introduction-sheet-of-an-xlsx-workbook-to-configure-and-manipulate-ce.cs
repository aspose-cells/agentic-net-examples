using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsIntroductionSheetDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (introduction sheet is assumed to be the first worksheet)
            Workbook workbook = new Workbook("Introduction.xlsx");

            // Access the first worksheet which contains the introduction content
            Worksheet introSheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Example 1: Determine the used range of the sheet
            // ------------------------------------------------------------
            Aspose.Cells.Range usedRange = introSheet.Cells.MaxDisplayRange;

            // If the sheet is empty, MaxDisplayRange may be null; handle gracefully.
            if (usedRange != null)
            {
                Console.WriteLine($"Used range: {usedRange.FirstRow}-{usedRange.FirstRow + usedRange.RowCount - 1}, " +
                                  $"{usedRange.FirstColumn}-{usedRange.FirstColumn + usedRange.ColumnCount - 1}");
            }
            else
            {
                Console.WriteLine("The worksheet is empty.");
            }

            // ------------------------------------------------------------
            // Example 2: Create a named range for the first 5 rows and 3 columns
            // ------------------------------------------------------------
            Aspose.Cells.Range headerRange = introSheet.Cells.CreateRange(0, 0, 5, 3);
            headerRange.Name = "IntroHeader";

            // Apply a style to make the header stand out
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style using a StyleFlag to affect all style attributes
            StyleFlag flag = new StyleFlag { All = true };
            headerRange.ApplyStyle(headerStyle, flag);

            // ------------------------------------------------------------
            // Example 3: Clear contents of a specific area (e.g., rows 10-12, columns 2-4)
            // ------------------------------------------------------------
            Aspose.Cells.Range clearRange = introSheet.Cells.CreateRange(9, 1, 3, 3); // zero‑based indices
            clearRange.ClearContents();

            // ------------------------------------------------------------
            // Example 4: Add a hyperlink to cell B2 within the introduction sheet
            // ------------------------------------------------------------
            introSheet.Hyperlinks.Add("B2", 1, 1, "https://www.aspose.com");

            // ------------------------------------------------------------
            // Save the modified workbook using the provided Save method
            // ------------------------------------------------------------
            workbook.Save("Introduction_Modified.xlsx");
        }
    }
}