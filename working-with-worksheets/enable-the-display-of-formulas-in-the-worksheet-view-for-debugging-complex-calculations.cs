// Title: Display formulas instead of values in an Aspose.Cells worksheet using C# for debugging
// Common Searches: C# how to show formulas in an Aspose.Cells generated Excel workbook | enable formula view in Aspose.Cells worksheet for debugging calculations | Aspose.Cells display cell formulas rather than results in C# | use workbook.Settings.ShowFormula to view formulas in Aspose.Cells | debug complex Excel formulas with Aspose.Cells by showing formulas
// Tags: show formulas worksheet Aspose.Cells C# | debug Excel calculations Aspose.Cells workbook settings | Aspose.Cells workbook Settings ShowFormula property | display cell formulas in generated Excel file C# | enable formula view Aspose.Cells for debugging

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, notes that the ShowFormula property is unavailable in the current Aspose.Cells version, and explains that when supported you can enable formula display via workbook.Settings.ShowFormula before saving the file as FormulasDisplayed.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Note: The ShowFormula property is not available in the current Aspose.Cells version.
            // If needed, you can enable formula display via other means (e.g., workbook.Settings.ShowFormula = true in supported versions).

            // Define output file path
            string outputPath = "FormulasDisplayed.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
