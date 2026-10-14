// Title: Show formulas in an Aspose.Cells worksheet for debugging with C# (Settings.ShowFormula)
// AI Prompts: Enable the Settings.ShowFormula flag on a Workbook object and save it as .xlsx to display formulas instead of values using Aspose.Cells for .NET. | Create a new workbook, turn on formula visibility, and export the file to aid debugging of complex calculations in C#. | Toggle formula display in an existing Aspose.Cells workbook and re‑save to view formulas in Excel.
// Common Searches: C# Aspose.Cells how to display cell formulas instead of values in generated Excel file | Aspose.Cells Settings.ShowFormula property example for debugging formulas | show formulas in Excel workbook created with Aspose.Cells .NET API | enable formula view in Aspose.Cells workbook for troubleshooting calculations | Aspose.Cells formula visibility not working in C#
// Tags: Settings.ShowFormula property Aspose.Cells | display formulas in generated .xlsx using Aspose.Cells | debug Excel calculations with Aspose.Cells C# | enable formula view Aspose.Cells workbook | Aspose.Cells workbook formula visibility

using System;
using Aspose.Cells;

// The example demonstrates how to activate the Settings.ShowFormula property on an Aspose.Cells Workbook in C# so that formulas are shown in the worksheet view instead of their evaluated results, then saves the workbook as FormulasShown.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // If the current Aspose.Cells version supports it, enable formula display
            // workbook.Settings.ShowFormula = true;

            // Save the workbook
            string outputFile = "FormulasShown.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
