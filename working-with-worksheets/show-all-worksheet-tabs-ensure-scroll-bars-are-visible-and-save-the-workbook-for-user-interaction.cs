// Title: C# – Show Worksheet Tabs and Scrollbars in Aspose.Cells and Save the Workbook
// Description: Creates a new Workbook, enables worksheet tabs, makes vertical and horizontal scrollbars visible via Workbook.Settings, and saves the file as an XLSX document ready for end‑user interaction.
// Keywords: Aspose.Cells ShowTabs | Aspose.Cells scrollbars | C# display worksheet tabs | enable vertical scrollbar Aspose | enable horizontal scrollbar Aspose | save workbook with UI settings | Aspose.Cells workbook UI configuration
// Common Searches: Aspose.Cells show worksheet tabs C# | how to enable scrollbars in Aspose.Cells workbook | C# Aspose.Cells set ShowTabs true | make vertical and horizontal scrollbars visible Aspose.Cells | save Excel file with tabs and scrollbars using Aspose
// Developer Intent: Enable worksheet tabs and both scrollbars in a generated Excel file and persist the workbook for user viewing.
// Use Cases: Programmatically prepare an Excel template where users see tabs and scrollbars immediately on open. | Generate reports that require full navigation controls without manual Excel configuration. | Distribute workbooks from a web service or desktop app with UI settings pre‑configured for better usability.
// AI Prompts: Generate C# code with Aspose.Cells that turns on ShowTabs, IsVScrollBarVisible, and IsHScrollBarVisible, then saves the workbook as XLSX. | Explain how Workbook.Settings properties affect the Excel UI and provide a sample implementation. | Create a reusable method that accepts a file path, configures tab and scrollbar visibility, and writes the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a new Workbook, enables worksheet tabs, makes vertical and horizontal scrollbars visible via Workbook.Settings, and saves the file as an XLSX document ready for end‑user interaction.
    class ShowTabsAndScrollbars
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Make sure worksheet tabs are displayed
            workbook.Settings.ShowTabs = true;

            // Ensure both vertical and horizontal scrollbars are visible
            workbook.Settings.IsVScrollBarVisible = true;
            workbook.Settings.IsHScrollBarVisible = true;

            // Save the workbook so the user can open it
            workbook.Save("WorkbookWithTabsAndScrollbars.xlsx", SaveFormat.Xlsx);
        }
    }
}
