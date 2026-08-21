// Title: Reset Worksheet TabId Sequentially with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to renumber the TabId of every worksheet in an Aspose.Cells workbook so that the IDs start at 1 and increase by one for each sheet. The code creates a workbook, adds sample sheets, reassigns TabId values, saves the file, and prints the IDs to confirm the sequential order.
// Keywords: Aspose.Cells TabId | reset worksheet TabId | sequential worksheet IDs | C# Aspose.Cells example | Excel tab order programmatically | Aspose.Cells workbook manipulation | standardize sheet TabId
// Common Searches: how to reset worksheet TabId Aspose.Cells | set Excel sheet TabId sequentially C# | Aspose.Cells change worksheet tab identifiers | renumber worksheet TabId after adding sheets | verify TabId values in saved workbook Aspose.Cells
// Developer Intent: Assign sequential TabId numbers to all worksheets in a workbook.
// Use Cases: Ensure a predictable tab order before exporting a workbook to another system. | Re‑assign TabId values after dynamically adding or removing sheets to keep IDs contiguous. | Validate TabId sequence when loading a workbook for automated processing or reporting.
// AI Prompts: Generate C# code using Aspose.Cells that resets each worksheet's TabId to start at 1 and saves the workbook. | Create a method that takes an existing Workbook object and reassigns sequential TabId values to its worksheets. | Explain how to read and verify TabId values after saving a workbook with Aspose.Cells, and why sequential IDs may be required.

using System;
using System.IO;
using Aspose.Cells;

namespace ExampleNamespace
{
    // This example demonstrates how to renumber the TabId of every worksheet in an Aspose.Cells workbook so that the IDs start at 1 and increase by one for each sheet. The code creates a workbook, adds sample sheets, reassigns TabId values, saves the file, and prints the IDs to confirm the sequential order.
    public class ResetWorksheetTabIds
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample worksheets
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Reset TabId sequentially starting from 1
                int nextTabId = 1;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.TabId = nextTabId;
                    nextTabId++;
                }

                // Define output file path
                string outputPath = "ResetTabIds.xlsx";

                // Save the workbook
                workbook.Save(outputPath);

                // Verify TabIds by loading the saved workbook if file exists
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    foreach (Worksheet sheet in loadedWorkbook.Worksheets)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" has TabId: {sheet.TabId}");
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to create file: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResetWorksheetTabIds.Run();
        }
    }
}
