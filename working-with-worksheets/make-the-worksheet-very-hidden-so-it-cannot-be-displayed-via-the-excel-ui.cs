// Title: Create a VeryHidden worksheet with Aspose.Cells for .NET (cannot be shown in Excel UI)
// Description: Demonstrates how to add a worksheet, rename it, set its VisibilityType to VeryHidden, save the workbook, and reload it to confirm the sheet remains hidden using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | VeryHidden worksheet | VisibilityType.VeryHidden | hide Excel sheet C# | protect hidden sheet .NET | programmatically hide worksheet | Excel very hidden sheet
// Common Searches: Aspose.Cells set worksheet to VeryHidden | How to hide a worksheet from Excel UI using Aspose.Cells | C# create very hidden sheet in Excel | Verify VeryHidden worksheet after saving Aspose.Cells | Unhide VeryHidden sheet programmatically Aspose
// Developer Intent: Hide a worksheet so it cannot be displayed through Excel’s UI by setting its VisibilityType to VeryHidden.
// Use Cases: Store sensitive data that should not be visible to end users | Keep configuration or metadata sheets hidden while still usable by formulas | Prevent accidental editing of intermediate calculation sheets | Embed macro‑related data without exposing it to users
// AI Prompts: Write C# code with Aspose.Cells to add multiple VeryHidden worksheets and save the workbook. | Show how to switch a worksheet between Visible, Hidden, and VeryHidden at runtime using Aspose.Cells. | Explain how to programmatically unhide a VeryHidden sheet for internal processing while keeping it hidden from end users.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a worksheet, rename it, set its VisibilityType to VeryHidden, save the workbook, and reload it to confirm the sheet remains hidden using Aspose.Cells for .NET.
    public class VeryHiddenWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default worksheet is added automatically)
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SecretSheet";

                // Make the worksheet very hidden so it cannot be shown via Excel UI
                sheet.VisibilityType = VisibilityType.VeryHidden;

                // Define output file path
                string outputPath = "VeryHiddenWorksheet.xlsx";

                // Save the workbook
                workbook.Save(outputPath);

                // Optional: reload to verify the visibility setting
                if (File.Exists(outputPath))
                {
                    Workbook loaded = new Workbook(outputPath);
                    Console.WriteLine("Worksheet visibility after reload: " + loaded.Worksheets[0].VisibilityType);
                }
                else
                {
                    Console.WriteLine("Failed to create the output file.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VeryHiddenWorksheetDemo.Run();
        }
    }
}
