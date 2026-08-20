// Title: Clone a PivotTable to a new worksheet with layout and formatting using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, extracts the first PivotTable, adds a new worksheet, clones the PivotTable with PivotTables.Add (preserving layout, styles, and filters), refreshes and calculates the copy, then saves the workbook.
// Keywords: Aspose.Cells C# clone pivot table | copy pivot table worksheet Aspose | preserve pivot formatting .NET | PivotTables.Add source pivot | refresh cloned pivot Aspose.Cells | duplicate pivot table programmatically | Aspose.Cells PivotTable API
// Common Searches: Aspose.Cells how to duplicate a PivotTable | C# copy PivotTable to another sheet Aspose | clone PivotTable layout formatting Aspose.Cells | refresh data after cloning PivotTable Aspose | programmatically add PivotTable copy in .NET
// Developer Intent: Programmatically create an exact copy of an existing PivotTable on a separate worksheet while retaining its design, filters, and calculations.
// Use Cases: Generate a reporting workbook that keeps the original data sheet untouched and adds a cloned PivotTable on a presentation sheet. | Automate workbook versioning where the source PivotTable remains static and a duplicated version is used for scenario analysis. | Build a client‑specific template that inserts a cloned PivotTable into a newly added worksheet for each processed file.
// AI Prompts: Show me C# code to clone a PivotTable to a new worksheet with Aspose.Cells, keeping layout and formatting intact. | How can I duplicate a PivotTable, refresh its data, and save the workbook using Aspose.Cells for .NET? | Explain the steps to copy a specific PivotTable when a worksheet contains multiple PivotTables in Aspose.Cells. | Provide an example of using PivotTables.Add to preserve styles and filters while cloning a PivotTable.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotClone
{
    // Loads a workbook, extracts the first PivotTable, adds a new worksheet, clones the PivotTable with PivotTables.Add (preserving layout, styles, and filters), refreshes and calculates the copy, then saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";
                const string resultPath = "ClonedPivotWorkbook.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                    return;
                }

                // Load the source workbook that contains the original PivotTable
                Workbook workbook = new Workbook(sourcePath);

                // Get the first worksheet (assumed to contain the original PivotTable)
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Ensure the worksheet actually has a PivotTable
                if (sourceSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTable found in the source worksheet.");
                    return;
                }

                // Retrieve the first PivotTable from the source worksheet
                PivotTable sourcePivot = sourceSheet.PivotTables[0];

                // Add a new worksheet that will host the cloned PivotTable
                Worksheet clonedSheet = workbook.Worksheets.Add("ClonedPivot");

                // Clone the PivotTable to the new worksheet.
                // The Add method copies layout and formatting from the source PivotTable.
                int clonedPivotIndex = clonedSheet.PivotTables.Add(sourcePivot, "A1", "ClonedPivotTable");

                // Refresh and calculate the cloned PivotTable to ensure data is up‑to‑date
                PivotTable clonedPivot = clonedSheet.PivotTables[clonedPivotIndex];
                clonedPivot.RefreshData();
                clonedPivot.CalculateData();

                // Save the workbook with the cloned PivotTable
                workbook.Save(resultPath);
                Console.WriteLine($"Cloned PivotTable saved to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
