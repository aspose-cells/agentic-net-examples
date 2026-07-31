// Title: Hide First Worksheet Tab, Show Formulas on Second Sheet, and Save Workbook – Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, adds a second sheet, hides the first sheet tab, enables formula view on the second sheet, and saves the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | hide worksheet tab | ShowFormulas | worksheet visibility | save workbook | Excel automation | IsVisible property | display formulas | XLSX export
// Common Searches: Aspose.Cells hide first worksheet tab C# | Show formulas on a specific sheet Aspose.Cells | Save workbook after changing sheet visibility Aspose.Cells | C# code to hide sheet tab and display formulas | How to use ShowFormulas property Aspose.Cells
// Developer Intent: Hide the first sheet tab, display formulas on the second sheet, and persist the workbook.
// Use Cases: Prepare a confidential summary sheet while exposing calculation formulas for auditors. | Generate an Excel template where users cannot navigate to the hidden sheet but can review raw formulas on another sheet. | Automate report creation that hides navigation tabs and shows formulas for developers during debugging.
// AI Prompts: Generate C# Aspose.Cells code that hides the first worksheet, enables ShowFormulas on the second worksheet, and saves the workbook as an XLSX file. | Explain how the IsVisible and ShowFormulas properties differ and when to use each in Aspose.Cells. | Provide a C# example that toggles worksheet visibility and formula display based on a configuration flag using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Creates a new workbook, adds a second sheet, hides the first sheet tab, enables formula view on the second sheet, and saves the file as XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with the default first worksheet
            Workbook workbook = new Workbook();

            // Add a second worksheet (named "Sheet2")
            workbook.Worksheets.Add("Sheet2");

            // Hide the first worksheet tab
            // Option 1: directly set the visibility flag
            workbook.Worksheets[0].IsVisible = false;
            // Optionally, you could use SetVisible(false, true) as well:
            // workbook.Worksheets[0].SetVisible(false, true);

            // Display formulas (instead of calculated results) on the second worksheet
            workbook.Worksheets[1].ShowFormulas = true;

            // Save the workbook with the applied changes
            workbook.Save("HiddenFirstSheet_ShowFormulasSecondSheet.xlsx", SaveFormat.Xlsx);
        }
    }
}
