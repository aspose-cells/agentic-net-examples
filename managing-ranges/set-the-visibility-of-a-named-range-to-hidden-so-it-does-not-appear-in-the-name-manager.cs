// Title: Hide a Named Range from Excel Name Manager with Aspose.Cells for .NET
// Description: Demonstrates how to create a named range in a workbook, set its IsVisible property to false, and save the file so the range is hidden from Excel's Name Manager using Aspose.Cells for C#.
// Keywords: Aspose.Cells hide named range | C# hide Excel named range | IsVisible false Aspose.Cells | Excel Name Manager hidden range | Aspose.Cells .NET named range visibility | programmatically hide named range | Excel hidden name Aspose
// Common Searches: hide named range Aspose.Cells C# | set IsVisible false for Excel name manager | Aspose.Cells hide name manager entry | C# create hidden named range in Excel | how to hide a named range using Aspose.Cells
// Developer Intent: Hide a workbook's named range so it does not appear in Excel's Name Manager.
// Use Cases: Store internal calculation data in a hidden range to keep it invisible to end‑users. | Distribute Excel files while preventing users from editing or seeing specific named ranges. | Embed configuration values or metadata in a hidden range that server‑side processes can read.
// AI Prompts: Write C# code with Aspose.Cells that adds a named range and sets IsVisible to false. | Explain the effect of the Name.IsVisible property on the Excel Name Manager UI. | Show how to hide multiple named ranges in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a named range in a workbook, set its IsVisible property to false, and save the file so the range is hidden from Excel's Name Manager using Aspose.Cells for C#.
    public class HideNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is "Sheet1")
                Worksheet sheet = workbook.Worksheets[0];

                // Add a named range that refers to cells A1:B2
                int nameIndex = workbook.Worksheets.Names.Add("HiddenRange");
                Name hiddenName = workbook.Worksheets.Names[nameIndex];
                hiddenName.RefersTo = "=Sheet1!$A$1:$B$2";

                // Set the visibility of the named range to hidden
                hiddenName.IsVisible = false;

                // Save the workbook to a file
                workbook.Save("HiddenNamedRange.xlsx");
                Console.WriteLine("Workbook saved successfully as HiddenNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
