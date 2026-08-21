// Title: Hide a worksheet with VisibilityType.VeryHidden using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a visible sheet, set the original worksheet's VisibilityType to VeryHidden, and save the file so the sheet cannot be shown through Excel's UI.
// Keywords: Aspose.Cells VeryHidden worksheet | VisibilityType.VeryHidden C# | hide sheet Excel UI .NET | Aspose.Cells hide worksheet programmatically | VeryHidden sheet Aspose.Cells example
// Common Searches: Aspose.Cells set worksheet VeryHidden | C# hide Excel sheet so user cannot unhide | VisibilityType VeryHidden Aspose.Cells tutorial | make worksheet invisible in Excel using .NET | Aspose.Cells VeryHidden sheet example
// Developer Intent: Set a worksheet to VeryHidden so it is inaccessible via the Excel interface.
// Use Cases: Store internal configuration data that end users must not see. | Protect proprietary formulas by moving them to a VeryHidden sheet. | Maintain an audit log worksheet that remains hidden on workbook open.
// AI Prompts: Provide C# code that sets a worksheet's VisibilityType to VeryHidden with Aspose.Cells and saves the workbook. | Show how to hide multiple worksheets as VeryHidden in a single Aspose.Cells .NET project. | Explain how to toggle a worksheet between Visible and VeryHidden using Aspose.Cells for C#.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a visible sheet, set the original worksheet's VisibilityType to VeryHidden, and save the file so the sheet cannot be shown through Excel's UI.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add an additional visible worksheet to satisfy the requirement
            workbook.Worksheets.Add("VisibleSheet");

            // Make the original first worksheet VeryHidden (cannot be shown via Excel UI)
            Worksheet hiddenSheet = workbook.Worksheets[0];
            hiddenSheet.VisibilityType = VisibilityType.VeryHidden;

            // Save the workbook
            workbook.Save("VeryHiddenSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
