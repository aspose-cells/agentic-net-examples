// Title: Aspose.Cells for .NET: Set Tab Bar Width, Hide Fourth Worksheet, and Save Workbook
// Description: Shows how to create a workbook, add extra sheets, set the SheetTabBarWidth (in 1/1000 of the window width), hide the fourth worksheet with Worksheet.SetVisible, and export the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells .NET | SheetTabBarWidth | hide worksheet Aspose.Cells | Worksheet.SetVisible example | save workbook XLSX | Excel tab bar width programmatically | customize Excel UI Aspose | worksheet visibility .NET | Aspose.Cells code sample
// Common Searches: Aspose.Cells set tab bar width | How to hide a specific worksheet with Aspose.Cells | Save workbook after changing sheet visibility in .NET | Worksheet.SetVisible parameters explained | Customize Excel tab bar using Aspose.Cells
// Developer Intent: Adjust the workbook’s tab bar width, hide the fourth sheet, and persist the changes to an XLSX file.
// Use Cases: Match Excel UI layout to a corporate design by controlling tab bar width. | Keep internal or draft worksheets hidden while distributing the file to end users. | Apply visibility and UI settings before exporting a report or template.
// AI Prompts: Generate C# code with Aspose.Cells that sets SheetTabBarWidth to 800, hides the worksheet named "Confidential", and saves the workbook as "Report.xlsx". | Explain the second boolean argument of Worksheet.SetVisible in Aspose.Cells, including its impact on workbook saving and how to use it for permanent vs. temporary hiding.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Shows how to create a workbook, add extra sheets, set the SheetTabBarWidth (in 1/1000 of the window width), hide the fourth worksheet with Worksheet.SetVisible, and export the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default has one worksheet)
            Workbook workbook = new Workbook();

            // Add three more worksheets so we have at least four sheets (indices 0-3)
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4");

            // Set the width of the worksheet tab bar (value is in 1/1000 of window width)
            workbook.Settings.SheetTabBarWidth = 1000; // Example: full width

            // Hide the fourth worksheet (index 3). Use SetVisible to follow the available rule.
            workbook.Worksheets[3].SetVisible(false, true);

            // Save the workbook with the applied settings
            workbook.Save("Result.xlsx", SaveFormat.Xlsx);
        }
    }
}
