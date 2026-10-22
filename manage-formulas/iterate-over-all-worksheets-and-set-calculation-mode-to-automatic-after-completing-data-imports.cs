// Title: Recalculate All Formulas in an Excel Workbook After Data Import Using Aspose.Cells for .NET
// Common Searches: how to recalculate all formulas in a workbook after importing data with Aspose.Cells C# | Aspose.Cells C# example to load, calculate formulas and save Excel file | refresh formulas in Excel after data import using Aspose.Cells .NET library | C# code to recalculate workbook formulas and export updated file with Aspose.Cells
// Tags: recalculate formulas Aspose.Cells C# | load workbook calculate save Aspose.Cells .NET | Excel data import formula refresh C# | Aspose.Cells automatic calculation mode setting | update workbook after data import Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing Excel workbook (input.xlsx) with Aspose.Cells, recalculates all formulas to reflect any imported data, and saves the updated workbook to output.xlsx while handling missing file and general exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing imported data
            var workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
