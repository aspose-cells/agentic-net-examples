// Title: Set Worksheet TabId Sequentially and Save Workbook – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add multiple sheets, assign each worksheet a TabId equal to its position (index + 1) to control tab order, save the file, and reload it to confirm the TabId values using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | set TabId programmatically | Excel tab order | save workbook Aspose | reload workbook verification | Excel automation | API example
// Common Searches: how to set TabId for worksheets in Aspose.Cells C# | Aspose.Cells assign tab order .NET | save workbook after changing TabId | verify TabId values after saving Excel file | C# example for worksheet TabId property
// Developer Intent: Assign a 1‑based TabId to every worksheet and persist the ordering by saving the workbook.
// Use Cases: Create a new workbook, add several sheets, set TabId = index + 1 to define the tab sequence, and save the workbook. | Update an existing workbook’s sheet order by recalculating TabId values and writing the changes back to disk. | After persisting, reload the workbook to ensure the TabId settings were stored correctly.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets, sets TabId to the sheet’s position plus one, and saves the workbook. | Write a reusable method that takes a Workbook object, reassigns TabId sequentially, and returns the saved file path. | Provide robust error‑handling for saving a workbook and confirming TabId values after reloading the file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Demonstrates how to create a workbook, add multiple sheets, assign each worksheet a TabId equal to its position (index + 1) to control tab order, save the file, and reload it to confirm the TabId values using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Rename the default worksheet to avoid name conflicts
                workbook.Worksheets[0].Name = "Sheet0";

                // Add additional worksheets for demonstration with unique names
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Set each worksheet's TabId to its zero‑based index plus one (TabId is 1‑based)
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    sheet.TabId = i + 1;
                }

                // Save the workbook to persist the TabId changes
                string outputPath = "WorkbookWithTabIds.xlsx";
                workbook.Save(outputPath);

                // Verify the file exists before loading
                if (File.Exists(outputPath))
                {
                    // Load the saved workbook to verify TabId values
                    Workbook loaded = new Workbook(outputPath);
                    for (int i = 0; i < loaded.Worksheets.Count; i++)
                    {
                        Console.WriteLine($"Worksheet '{loaded.Worksheets[i].Name}' TabId: {loaded.Worksheets[i].TabId}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
