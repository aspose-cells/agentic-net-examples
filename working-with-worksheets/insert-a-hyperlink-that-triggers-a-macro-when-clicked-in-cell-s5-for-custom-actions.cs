// Title: Insert a Macro‑Triggering Hyperlink in Cell S5 with Aspose.Cells for .NET (C#)
// Description: Shows how to create a new workbook, enable macros, and add a hyperlink in cell S5 that runs the macro “MyMacro”. The example sets the link’s display text and screen tip, then saves the file as MacroHyperlink.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | hyperlink to macro | enable macros | cell S5 | run macro from hyperlink | Excel automation | MacroHyperlink.xlsx
// Common Searches: Aspose.Cells add hyperlink that runs a macro | C# create macro hyperlink in Excel | How to enable macros in Aspose.Cells workbook | Set screen tip for hyperlink Aspose.Cells | Add clickable macro link to specific cell using Aspose
// Developer Intent: Generate an Excel file where clicking cell S5 launches the macro MyMacro.
// Use Cases: Add a button‑like link in a generated report that refreshes data via a custom macro. | Provide end‑users with a one‑click option to apply formatting macros in a template worksheet. | Create a dashboard hyperlink that triggers a macro to export data to another format.
// AI Prompts: Write C# code with Aspose.Cells to add a hyperlink in cell S5 that calls a macro named MyMacro, including enabling macros and setting a screen tip. | Explain how to modify the hyperlink to reference a macro stored in a different workbook using Aspose.Cells. | Generate C# code that adds multiple macro‑triggering hyperlinks across a range of cells, each with custom display text and screen tip, using Aspose.Cells for .NET.

using Aspose.Cells;

// Shows how to create a new workbook, enable macros, and add a hyperlink in cell S5 that runs the macro “MyMacro”. The example sets the link’s display text and screen tip, then saves the file as MacroHyperlink.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable macros in the workbook
        workbook.Settings.EnableMacros = true;

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell S5 that points to a macro named "MyMacro"
        int hyperlinkIndex = worksheet.Hyperlinks.Add("S5", 1, 1, "MyMacro");

        // Set the display text and screen tip for the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Run My Macro";
        worksheet.Hyperlinks[hyperlinkIndex].ScreenTip = "Click to execute the macro";

        // Save the workbook
        workbook.Save("MacroHyperlink.xlsx");
    }
}
