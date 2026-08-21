// Title: Aspose.Cells .NET: Set Worksheet VeryHidden, Freeze Panes, Then Make Visible
// Description: Demonstrates how to hide a worksheet as VeryHidden, apply FreezePanes at cell C3 while hidden, restore visibility, and optionally remove a temporary sheet, using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VeryHidden worksheet | freeze panes hidden sheet Aspose.Cells | make worksheet visible Aspose.Cells | temporary worksheet Aspose.Cells | C# Aspose.Cells hide sheet | Aspose.Cells workbook save
// Common Searches: Aspose.Cells set worksheet VeryHidden and freeze panes | freeze panes on a hidden worksheet using Aspose.Cells | unhide VeryHidden sheet after applying FreezePanes Aspose.Cells | remove temporary sheet after changing visibility Aspose.Cells | C# Aspose.Cells hide sheet then show
// Developer Intent: Hide a worksheet as VeryHidden, freeze panes while hidden, then reveal the sheet.
// Use Cases: Secure a sheet by hiding it during layout configuration such as freeze panes. | Satisfy Aspose.Cells requirement for at least one visible sheet before applying VeryHidden. | Programmatically clean up a temporary worksheet after visibility and freeze settings are applied.
// AI Prompts: Write C# code with Aspose.Cells to set a worksheet to VeryHidden, freeze panes at D4, then make it visible and delete a temporary sheet. | Explain why Aspose.Cells needs a visible worksheet before another can be set to VeryHidden and show the proper workaround. | Provide robust error handling for freezing panes on a hidden worksheet using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityAndFreezeDemo
{
    // Demonstrates how to hide a worksheet as VeryHidden, apply FreezePanes at cell C3 while hidden, restore visibility, and optionally remove a temporary sheet, using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();

                // Reference the first (target) worksheet
                Worksheet targetSheet = workbook.Worksheets[0];

                // Add a temporary visible worksheet so the workbook always has at least one visible sheet
                Worksheet tempSheet = workbook.Worksheets.Add("Temp");

                // Hide the target worksheet as VeryHidden (requires another visible sheet)
                targetSheet.VisibilityType = VisibilityType.VeryHidden;

                // Freeze panes while the sheet is hidden (freeze at cell C3, 3 rows and 3 columns)
                targetSheet.FreezePanes("C3", 3, 3);

                // Make the target worksheet visible again
                targetSheet.VisibilityType = VisibilityType.Visible;

                // Remove the temporary worksheet (optional)
                int tempIndex = workbook.Worksheets.IndexOf(tempSheet);
                if (tempIndex >= 0)
                {
                    workbook.Worksheets.RemoveAt(tempIndex);
                }

                // Save the workbook
                string outputPath = "VisibilityAndFreezeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
