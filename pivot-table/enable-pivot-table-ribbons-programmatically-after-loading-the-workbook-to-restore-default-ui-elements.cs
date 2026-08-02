// Title: Restore Default Pivot Table Ribbon and Enable Wizard/Field List with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, clears any custom RibbonXml, disables the HidePivotFieldList flag, iterates through all worksheets to set EnableWizard and EnableFieldList on each PivotTable, and saves the file so the standard Excel pivot UI appears when opened.
// Keywords: Aspose.Cells C# | pivot table ribbon reset | EnableWizard Aspose.Cells | EnableFieldList Aspose.Cells | RibbonXml clear | HidePivotFieldList false | restore Excel UI programmatically | pivot wizard activation | default pivot ribbon
// Common Searches: how to show pivot table ribbon after loading workbook Aspose.Cells | enable pivot wizard and field list with C# Aspose.Cells | reset RibbonXml in Excel file using Aspose.Cells | programmatically display pivot field list in .NET | Aspose.Cells restore default pivot UI
// Developer Intent: Programmatically re‑enable the Excel pivot‑table ribbon, wizard, and field list for all pivot tables in a workbook using Aspose.Cells for .NET.
// Use Cases: A workbook imported from another source hides the pivot ribbon; the code restores the default UI for end users. | Multiple worksheets contain pivot tables that need the wizard and field list enabled without manual editing. | Automated report generation pipelines require the pivot UI to be visible when the file is opened in Excel.
// AI Prompts: Generate C# code with Aspose.Cells that clears RibbonXml and makes the pivot field list visible for every pivot table. | Create a reusable method that accepts a file path, enables the pivot wizard and field list, and saves the workbook. | Show an example that processes a workbook containing several sheets and pivot tables, restoring the default pivot ribbon UI programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing workbook, clears any custom RibbonXml, disables the HidePivotFieldList flag, iterates through all worksheets to set EnableWizard and EnableFieldList on each PivotTable, and saves the file so the standard Excel pivot UI appears when opened.
class EnablePivotRibbonDemo
{
    static void Main()
    {
        // Load the workbook that contains pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Clear any custom Ribbon XML to restore the default Ribbon UI
        workbook.RibbonXml = null; // or string.Empty

        // Ensure the PivotTable field list is not hidden by workbook settings
        workbook.Settings.HidePivotFieldList = false;

        // Enable wizard and field list for each pivot table in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pivot in sheet.PivotTables)
            {
                pivot.EnableWizard = true;
                pivot.EnableFieldList = true;
            }
        }

        // Save the workbook with the restored UI elements
        workbook.Save("output.xlsx");
    }
}
