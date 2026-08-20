// Title: Retrieve Worksheet TabId Values with Aspose.Cells for .NET
// Description: Shows how to create or load a workbook, set custom TabId values, and read each worksheet's TabId using the Worksheet.TabId property in C#.
// Keywords: Aspose.Cells | Worksheet.TabId | C# | .NET | retrieve TabId | worksheet identifier | Aspose.Cells example | get sheet TabId | Workbook TabId property
// Common Searches: Aspose.Cells get worksheet TabId | Worksheet TabId C# example | How to read TabId property in Aspose.Cells | List all sheet TabIds Aspose.Cells | Retrieve TabId for each worksheet .NET
// Developer Intent: Obtain the TabId of every worksheet in an Aspose.Cells workbook.
// Use Cases: Log worksheet names with their TabId for debugging. | Map sheet names to TabId values to synchronize with external systems. | Validate custom TabId assignments before saving the file. | Generate a report of sheet identifiers for documentation.
// AI Prompts: Write C# code that extracts all worksheet TabId values from an Aspose.Cells workbook and stores them in a Dictionary<string, int>. | Create a script to export worksheet names and their TabId to a CSV file using Aspose.Cells. | Explain how to compare TabId collections of two workbooks to detect mismatched sheet identifiers.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Shows how to create or load a workbook, set custom TabId values, and read each worksheet's TabId using the Worksheet.TabId property in C#.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path) or create a new one.
            // Here we create a new workbook for demonstration.
            Workbook workbook = new Workbook();

            // Add a few worksheets to have multiple sheets.
            workbook.Worksheets.Add("FirstSheet");
            workbook.Worksheets.Add("SecondSheet");
            workbook.Worksheets.Add("ThirdSheet");

            // Optionally set custom TabId values to see distinct results.
            workbook.Worksheets[0].TabId = 101;
            workbook.Worksheets[1].TabId = 202;
            workbook.Worksheets[2].TabId = 303;

            // Iterate through each worksheet and retrieve its TabId.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet Name: {sheet.Name}, TabId: {sheet.TabId}");
            }

            // Save the workbook if needed.
            // workbook.Save("TabIdDemo.xlsx");
        }
    }
}
