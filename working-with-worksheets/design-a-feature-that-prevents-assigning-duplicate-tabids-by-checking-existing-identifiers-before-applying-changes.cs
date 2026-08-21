// Title: Prevent Duplicate Worksheet TabIds in Aspose.Cells for .NET
// Description: Demonstrates how to assign a unique TabId to each worksheet in an Aspose.Cells workbook. The sample creates a workbook, adds three sheets, checks existing TabIds before setting a new value, resolves conflicts by using the next highest integer, saves the file, reloads it, and prints the final TabIds.
// Keywords: Aspose.Cells TabId uniqueness | C# prevent duplicate worksheet TabId | assign unique TabId Aspose.Cells | check existing TabId before setting | resolve TabId conflict .NET | worksheet TabId validation | Aspose.Cells workbook TabId example
// Common Searches: how to ensure unique TabId for worksheets in Aspose.Cells | C# check duplicate TabId Aspose.Cells | assign incremental TabId when conflict occurs | Aspose.Cells get and set worksheet TabId | prevent duplicate sheet identifiers .NET
// Developer Intent: Guarantee that every worksheet in a workbook receives a distinct TabId by detecting existing IDs and automatically assigning a new one when a duplicate is found.
// Use Cases: Dynamically adding worksheets and needing guaranteed unique TabIds for UI navigation. | Validating and correcting TabIds after importing an external workbook to avoid identifier clashes. | Debugging workbook structure by listing each sheet’s TabId before publishing or processing.
// AI Prompts: Write a C# method for Aspose.Cells that assigns a unique TabId to a worksheet, checking the workbook for existing IDs and incrementing the maximum value on conflict. | Generate code that scans all worksheets in a workbook, detects duplicate TabIds, and resolves them by assigning sequential IDs. | Provide an example that logs duplicate TabId detection and automatically fixes it using Aspose.Cells for .NET.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to assign a unique TabId to each worksheet in an Aspose.Cells workbook. The sample creates a workbook, adds three sheets, checks existing TabIds before setting a new value, resolves conflicts by using the next highest integer, saves the file, reloads it, and prints the final TabIds.
    public class WorksheetTabIdUniqueDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            Worksheet sheet2 = workbook.Worksheets.Add("Second");
            Worksheet sheet3 = workbook.Worksheets.Add("Third");

            // Desired TabId to assign
            int desiredTabId = 107;

            // Assign a unique TabId to each worksheet
            AssignUniqueTabId(sheet1, desiredTabId, workbook);
            AssignUniqueTabId(sheet2, desiredTabId, workbook);
            AssignUniqueTabId(sheet3, desiredTabId, workbook);

            // Save the workbook to verify the TabIds
            string outputPath = "WorksheetTabIdUniqueDemo.xlsx";
            workbook.Save(outputPath);

            // Reload and display the TabIds if the file exists
            if (File.Exists(outputPath))
            {
                Workbook loadedWorkbook = new Workbook(outputPath);
                foreach (Worksheet ws in loadedWorkbook.Worksheets)
                {
                    Console.WriteLine($"Worksheet \"{ws.Name}\" TabId: {ws.TabId}");
                }
            }
            else
            {
                Console.WriteLine($"Failed to locate the saved file: {outputPath}");
            }
        }

        // Ensures the worksheet receives a TabId that is not already used in the workbook
        private static void AssignUniqueTabId(Worksheet targetSheet, int desiredId, Workbook workbook)
        {
            // Check if any worksheet already uses the desired TabId
            bool duplicateExists = workbook.Worksheets.Any(ws => ws != targetSheet && ws.TabId == desiredId);

            if (!duplicateExists)
            {
                // No conflict, assign the desired TabId
                targetSheet.TabId = desiredId;
            }
            else
            {
                // Find the maximum TabId currently used and assign the next integer
                int maxExistingId = workbook.Worksheets.Max(ws => ws.TabId);
                int newUniqueId = maxExistingId + 1;
                targetSheet.TabId = newUniqueId;
            }
        }
    }
}
