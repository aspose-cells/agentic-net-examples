// Title: How to set a custom row height and auto‑fit the next row using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, sets the height of row 0 to 20 points, writes long strings into cells A2 and B2, calls worksheet.AutoFitRow(1), and saves the file. | Write a C# snippet with Aspose.Cells that assigns a specific height to the first worksheet row and then automatically adjusts the height of the second row based on its cell contents.
// Common Searches: Aspose.Cells C# set row height then autofit next row example | C# Aspose.Cells how to auto‑fit a row after setting custom height for another row | programmatically adjust row height and auto‑fit rows in Excel using Aspose.Cells .NET | set specific row height and auto‑fit another row Aspose.Cells C# tutorial
// Tags: Aspose.Cells row height customization | AutoFitRow usage in .NET | Excel row auto‑fit based on content C# | worksheet row dimension API Aspose.Cells | set specific row height Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, sets the first row height to 20 points, fills cells A2 and B2 with long text, calls AutoFitRow(1) to resize the second row automatically, and saves the workbook as SetRowHeightAndAutoFitNextRow.xlsx.
    public class SetRowHeightAndAutoFitNextRow
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a custom height for the first row (index 0)
            worksheet.Cells.SetRowHeight(0, 20); // height in points

            // Populate the second row (index 1) with content that requires auto‑fit
            worksheet.Cells["A2"].PutValue("This is a long piece of text that will cause the row height to increase when auto‑fitted.");
            worksheet.Cells["B2"].PutValue("Additional long text in the same row.");

            // Auto‑fit the second row based on its content
            worksheet.AutoFitRow(1);

            // Save the workbook
            workbook.Save("SetRowHeightAndAutoFitNextRow.xlsx");
        }
    }
}
