// Title: Insert a macro‑triggering hyperlink in cell S5 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable macros, add a hyperlink to cell S5 that calls a VBA macro named "MyMacro", set custom display text and screen tip, and save the workbook using Aspose.Cells in C#.
// Keywords: Aspose.Cells C# hyperlink macro | add macro hyperlink Excel | enable macros Aspose.Cells | cell S5 hyperlink | Excel hyperlink screen tip | run VBA from hyperlink | Aspose.Cells workbook automation
// Common Searches: Aspose.Cells add hyperlink that runs a macro | C# create Excel hyperlink to VBA macro | how to enable macros in Aspose.Cells workbook | set hyperlink display text and tooltip in Excel using Aspose | insert macro link into specific cell with Aspose.Cells
// Developer Intent: Create a clickable link in cell S5 that launches the VBA macro MyMacro when the user clicks it.
// Use Cases: Provide end‑users a one‑click option to refresh data via a macro in generated reports. | Embed a setup macro link in a template workbook to configure custom settings on demand. | Add a tooltip‑enabled hyperlink that guides users to execute formatting or validation macros directly from the sheet.
// AI Prompts: Write C# code with Aspose.Cells that adds a hyperlink in cell S5 to execute a macro called 'MyMacro', including custom display text and a screen tip. | Explain the steps required to enable macros in an Aspose.Cells workbook before inserting a macro‑triggering hyperlink. | Show how to update an existing Excel file using Aspose.Cells to insert or replace a hyperlink that calls a specific VBA macro.

using Aspose.Cells;

// Demonstrates how to enable macros, add a hyperlink to cell S5 that calls a VBA macro named "MyMacro", set custom display text and screen tip, and save the workbook using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Enable macros in the workbook
        wb.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Add a hyperlink to cell S5 that points to a macro named "MyMacro"
        int hyperlinkIndex = ws.Hyperlinks.Add("S5", 1, 1, "MyMacro");

        // Set the display text and screen tip for the hyperlink
        ws.Hyperlinks[hyperlinkIndex].TextToDisplay = "Run Macro";
        ws.Hyperlinks[hyperlinkIndex].ScreenTip = "Click to execute MyMacro";

        // Save the workbook
        wb.Save("HyperlinkMacroDemo.xlsx");
    }
}
