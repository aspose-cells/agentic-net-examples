// Title: Get each worksheet's TabId using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook with Aspose.Cells, loops through all worksheets, reads the internal Worksheet.TabId property, and prints the sheet name with its TabId to the console.
// Keywords: Aspose.Cells TabId | Worksheet.TabId C# | retrieve worksheet TabId | Aspose.Cells get sheet TabId | C# Excel internal TabId
// Common Searches: Aspose.Cells how to read worksheet TabId | C# get TabId of Excel sheets | list TabId for all worksheets Aspose | Worksheet.TabId property example
// Developer Intent: Extract the TabId value for every worksheet in a loaded workbook.
// Use Cases: Log TabId values to verify sheet order after programmatic re‑ordering. | Create a name‑to‑TabId map for synchronizing custom metadata across workbooks. | Validate that expected TabId numbers exist after adding or deleting sheets.
// AI Prompts: Generate C# code that writes each worksheet's name and TabId to a CSV file using Aspose.Cells. | Provide a method returning Dictionary<string, int> where keys are worksheet names and values are their TabId values. | Explain the difference between a worksheet's TabId and its index, and show how TabId can be used to track sheets after they are moved.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Loads an Excel workbook with Aspose.Cells, loops through all worksheets, reads the internal Worksheet.TabId property, and prints the sheet name with its TabId to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Retrieve the internal TabId of the current worksheet
                int tabId = sheet.TabId;

                // Output the worksheet name and its TabId
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has TabId: {tabId}");
            }
        }
    }
}
