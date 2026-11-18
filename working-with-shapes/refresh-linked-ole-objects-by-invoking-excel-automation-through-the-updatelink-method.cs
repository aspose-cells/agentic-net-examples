// Title: Refresh OLE linked objects in an Excel workbook using Aspose.Cells CalculateFormula in C#
// Common Searches: C# Aspose.Cells refresh OLE links after loading workbook | Update external OLE references in .xlsx with Aspose.Cells CalculateFormula | Programmatically recalculate formulas and OLE objects using Aspose.Cells .NET | How to refresh embedded OLE objects in Excel file via Aspose.Cells | Aspose.Cells workbook.CalculateFormula to sync linked OLE data
// Tags: Aspose.Cells CalculateFormula OLE refresh | C# refresh OLE links Excel workbook | update external references Aspose.Cells | recalculate formulas Aspose.Cells .NET | programmatic OLE link update .xlsx

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, calls Workbook.CalculateFormula to recalculate all formulas and refresh external references—including OLE linked objects—and then saves the updated workbook, handling missing files and runtime errors gracefully.
class RefreshOleLinks
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Ensure the input file exists to prevent FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh formulas and external references (including OLE links)
            workbook.CalculateFormula();

            // Save the refreshed workbook
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook refreshed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Gracefully handle any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
