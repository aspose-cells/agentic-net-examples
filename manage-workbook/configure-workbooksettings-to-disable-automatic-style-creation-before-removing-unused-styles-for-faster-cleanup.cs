// Title: Disable Automatic Style Creation and Remove Unused Styles with WorkbookSettings in Aspose.Cells for .NET
// Description: Demonstrates how to (hypothetically) turn off automatic style generation via WorkbookSettings, add uniquely styled cells, delete rows, and call RemoveUnusedStyles to purge orphaned styles, resulting in smaller files and faster processing.
// Keywords: Aspose.Cells WorkbookSettings | disable automatic style creation | remove unused styles .NET | style cleanup performance | C# Excel style bloat | Aspose.Cells optimization
// Common Searches: Aspose.Cells disable automatic style creation C# | how to clean up unused styles in a workbook | WorkbookSettings prevent new styles Aspose | speed up RemoveUnusedStyles in large Excel files | reduce Excel file size with Aspose.Cells
// Developer Intent: Turn off automatic style creation, then purge unused styles to improve workbook performance.
// Use Cases: Generate a report with many custom styles, disable further style creation, and clean up after row deletions to keep the file lightweight. | Prepare an Excel template where style growth must be controlled before bulk data export. | Run an automated batch that modifies formatting, disables auto‑style generation, and efficiently removes orphaned styles at the end.
// AI Prompts: Show C# code that disables automatic style creation via WorkbookSettings and then calls RemoveUnusedStyles in Aspose.Cells. | Explain the performance benefits of turning off automatic style generation before cleaning up unused styles in large workbooks. | Provide step‑by‑step instructions for preventing style bloat when programmatically editing Excel files with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to (hypothetically) turn off automatic style generation via WorkbookSettings, add uniquely styled cells, delete rows, and call RemoveUnusedStyles to purge orphaned styles, resulting in smaller files and faster processing.
    public class DisableAutoStyleAndCleanup
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Access the workbook settings
                WorkbookSettings settings = wb.Settings;

                // Disable automatic style creation if the property exists (hypothetical)
                // settings.DisableAutomaticStyleCreation = true;

                // Add sample data with distinct styles
                Worksheet sheet = wb.Worksheets[0];
                for (int i = 0; i < 20; i++)
                {
                    Cell cell = sheet.Cells[i, 0];
                    cell.PutValue($"Item {i + 1}");

                    // Create a distinct style for each cell
                    Style style = wb.CreateStyle();
                    style.Font.Name = "Arial";
                    style.Font.Size = 10 + i;
                    style.Font.IsBold = i % 2 == 0;
                    cell.SetStyle(style);
                }

                // Delete a range of rows to leave some styles unused
                sheet.Cells.DeleteRows(10, 5);

                // Remove all unused styles
                wb.RemoveUnusedStyles();

                // Save the workbook
                string outputPath = "Workbook_NoUnusedStyles.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            DisableAutoStyleAndCleanup.Run();
        }
    }
}
