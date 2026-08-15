// Title: Aspose.Cells .NET – Apply a Custom Number Format to Show Negative Values in Red Parentheses
// Description: This example creates a workbook, inserts a positive and a negative amount, defines a custom style that formats negatives in red parentheses, applies the style to cells A1:A2 with a StyleFlag, saves the file, and confirms the format persists after reloading.
// Keywords: Aspose.Cells custom number format | negative numbers red parentheses C# | StyleFlag number format Aspose.Cells | Excel red parentheses formatting .NET | financial spreadsheet styling Aspose
// Common Searches: format negative numbers in red parentheses Aspose.Cells | custom number format for negatives C# Aspose | apply style to cell range Aspose.Cells .NET | how to use StyleFlag in Aspose.Cells
// Developer Intent: Implement a custom number format that displays negative figures in red parentheses across a cell range using Aspose.Cells for .NET.
// Use Cases: Generate financial reports where losses appear in red parentheses for quick visual identification. | Standardize negative‑value styling across multiple worksheets without manual cell editing. | Programmatically verify that custom number formats are retained after saving and reopening an Excel file.
// AI Prompts: Write C# code with Aspose.Cells that formats a range so negative numbers appear in red parentheses. | Show how to create a StyleFlag that only changes the number format and apply it to cells A1:A2. | Explain the steps to confirm that a custom number format is saved correctly after exporting and reloading the workbook.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // This example creates a workbook, inserts a positive and a negative amount, defines a custom style that formats negatives in red parentheses, applies the style to cells A1:A2 with a StyleFlag, saves the file, and confirms the format persists after reloading.
    public class NegativeNumberRedParenthesesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a positive and a negative value in cells A1 and A2
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(-1234.56);

            // Create a custom style with a number format that shows negatives in red parentheses
            Style style = workbook.CreateStyle();
            // Format: positive numbers as normal, negatives in red with parentheses
            style.Custom = "_-€ #,##0.00;[Red]_-€ -#,##0.00";

            // Use StyleFlag to apply only the number format part of the style
            StyleFlag styleFlag = new StyleFlag
            {
                NumberFormat = true
            };

            // Apply the style to the range A1:A2
            AsposeRange range = sheet.Cells.CreateRange("A1", "A2");
            range.ApplyStyle(style, styleFlag);

            // Save the workbook to a file
            string filePath = "NegativeNumberRedParenthesesDemo.xlsx";
            workbook.Save(filePath);

            // Optional: reload the workbook to verify the custom format was saved
            if (File.Exists(filePath))
            {
                Workbook verifyWorkbook = new Workbook(filePath);
                Worksheet verifySheet = verifyWorkbook.Worksheets[0];
                Console.WriteLine("Cell A1 format: " + verifySheet.Cells["A1"].GetStyle().Custom);
                Console.WriteLine("Cell A2 format: " + verifySheet.Cells["A2"].GetStyle().Custom);
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
    }
}
