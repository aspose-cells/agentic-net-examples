// Title: Hide a Named Range in Aspose.Cells for .NET (C#) – exclude it from the Name Manager
// Description: Demonstrates how to create a workbook, add a named range, set its IsVisible property to false, verify the setting, and save the file so the name is hidden from Excel's Name Manager.
// Keywords: Aspose.Cells hide named range | C# set named range visibility | Name.IsVisible Aspose.Cells | hide name manager entry | Aspose.Cells .NET hide range | Excel hidden named range programmatically
// Common Searches: how to hide a named range using Aspose.Cells C# | Aspose.Cells set IsVisible false for name | remove name from Name Manager Aspose.Cells .NET | programmatically hide Excel named range C# | Aspose.Cells hide internal range
// Developer Intent: Hide a defined name so it is not listed in Excel's Name Manager while remaining usable in formulas.
// Use Cases: Store helper ranges that users shouldn't see but formulas can reference. | Keep intermediate calculation data out of the UI for a cleaner worksheet. | Conceal sensitive cell references to protect workbook structure.
// AI Prompts: Write C# code with Aspose.Cells that hides every named range whose name starts with "Temp_". | Show how to toggle the IsVisible property of a specific named range at runtime. | Provide a snippet that lists all hidden named ranges in an existing workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a named range, set its IsVisible property to false, verify the setting, and save the file so the name is hidden from Excel's Name Manager.
    public class HideNamedRangeDemo
    {
        // Entry point for the application
        public static void Main()
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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a named range (e.g., cells A1:B2)
            int nameIndex = workbook.Worksheets.Names.Add("HiddenRange");
            Name hiddenName = workbook.Worksheets.Names[nameIndex];
            hiddenName.RefersTo = "=Sheet1!$A$1:$B$2";

            // Hide the named range so it won't appear in the Name Manager
            hiddenName.IsVisible = false;

            // Verify the visibility setting
            Console.WriteLine($"Is '{hiddenName.Text}' visible? {hiddenName.IsVisible}");

            // Save the workbook
            workbook.Save("HiddenNamedRange.xlsx");
        }
    }
}
