// Title: Display worksheet tabs, enable scrollbars, and save an Excel file with Aspose.Cells (C#)
// Description: Creates a new Workbook, sets ShowTabs, IsVScrollBarVisible, and IsHScrollBarVisible to true, writes sample data to cell A1, and saves the file as XLSX for end‑user interaction.
// Keywords: Aspose.Cells C# show tabs | Aspose.Cells enable scrollbars | Aspose.Cells workbook settings | save Excel file Aspose.Cells | display tabs and scrollbars | C# Excel generation Aspose
// Common Searches: How to make sheet tabs visible using Aspose.Cells C# | Enable vertical scrollbar in generated Excel with Aspose.Cells | Enable horizontal scrollbar in generated Excel with Aspose.Cells | Saving an Excel workbook with default UI elements via Aspose.Cells | C# Aspose.Cells workbook Settings for tabs and scrollbars
// Developer Intent: Configure a workbook to display sheet tabs and both scrollbars, then save it for user interaction.
// Use Cases: Prepare a user‑friendly template where sheet navigation is obvious. | Generate reports that retain Excel UI elements for easier scrolling. | Distribute Excel files that open with the default view (tabs and scrollbars) intact.
// AI Prompts: Generate C# Aspose.Cells code that toggles worksheet tabs and scrollbars based on a boolean flag. | Explain how Workbook.Settings properties influence the Excel UI when the file is opened. | Provide a step‑by‑step guide to configure and save a workbook with visible tabs and scrollbars in .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsSettingsDemo
{
    // Creates a new Workbook, sets ShowTabs, IsVScrollBarVisible, and IsHScrollBarVisible to true, writes sample data to cell A1, and saves the file as XLSX for end‑user interaction.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Ensure worksheet tabs are displayed
            workbook.Settings.ShowTabs = true;

            // Ensure vertical and horizontal scroll bars are visible
            workbook.Settings.IsVScrollBarVisible = true;
            workbook.Settings.IsHScrollBarVisible = true;

            // Optionally add some data to demonstrate the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Demo data");

            // Save the workbook for user interaction (lifecycle save rule)
            workbook.Save("WorkbookWithTabsAndScrollbars.xlsx", SaveFormat.Xlsx);
        }
    }
}
