// Title: Set Worksheet TabId Sequentially and Save Workbook with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds extra sheets, sets each worksheet's TabId to its zero‑based Index + 1, prints the assigned IDs, and saves the file so the tab order is persisted.
// Keywords: Aspose.Cells TabId C# | set worksheet TabId | worksheet tab order Aspose | Aspose.Cells save workbook | C# Excel tab ordering | Worksheet.Index property | Excel TabId property
// Common Searches: Aspose.Cells set TabId C# | How to change worksheet tab order programmatically | Worksheet.TabId property example | Save workbook after modifying TabId Aspose | C# set Excel sheet TabId Aspose.Cells
// Developer Intent: Assign each worksheet a TabId equal to its index + 1 and save the workbook to make the tab order permanent.
// Use Cases: Generate multi‑sheet reports where the visual tab sequence must match processing logic. | Prepare workbooks for downstream systems that reference sheets by TabId rather than by name. | Integrate with UI components that display worksheet tabs based on TabId values.
// AI Prompts: Write C# code using Aspose.Cells to iterate through all worksheets, set TabId = Index + 1, and save the workbook as XLSX. | Explain the difference between Worksheet.Index and Worksheet.TabId in Aspose.Cells and why a Save call is required after changing TabId. | Add robust error handling and logging for setting TabId and saving the workbook, including output of each assigned TabId. | Create a unit test that verifies TabId values are sequential after running the SetWorksheetTabIds example.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds extra sheets, sets each worksheet's TabId to its zero‑based Index + 1, prints the assigned IDs, and saves the file so the tab order is persisted.
    public class SetWorksheetTabIds
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Add additional worksheets for demonstration
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Iterate through all worksheets and set TabId = Index + 1
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Worksheet.Index is zero‑based, TabId expects a positive identifier
                    sheet.TabId = sheet.Index + 1;
                    // Optional: display the assigned TabId
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" assigned TabId: {sheet.TabId}");
                }

                // Save the workbook to apply the TabId changes
                string outputPath = "WorkbookWithTabIds.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
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
            SetWorksheetTabIds.Run();
        }
    }
}
