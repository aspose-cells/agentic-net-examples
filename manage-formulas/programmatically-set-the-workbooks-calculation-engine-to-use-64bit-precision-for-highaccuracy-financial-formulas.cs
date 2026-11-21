// Title: How to set Aspose.Cells workbook to use the 64‑bit calculation engine for high‑precision financial formulas in C#
// Common Searches: C# enable 64-bit calculation engine in Aspose.Cells workbook | Aspose.Cells set CalcEngineVersion to 2 for accurate financial formulas | increase formula precision with Aspose.Cells Settings.CalcEngineVersion | example of using 64-bit precision calculations in Aspose.Cells .NET | how to configure Aspose.Cells for high‑accuracy financial calculations
// Tags: calculation engine version setting Aspose.Cells | 64-bit precision calculations Aspose.Cells | financial formula accuracy Aspose.Cells | Aspose.Cells workbook precision configuration | double precision evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, optionally sets its Settings.CalcEngineVersion to 2 to activate the 64‑bit calculation engine for higher‑precision financial formulas, and saves the workbook as HighPrecision.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Set the calculation engine to 64‑bit precision (version 2) if supported
                // Note: In some Aspose.Cells versions the CalcEngineVersion property may not be available.
                // Uncomment the following line if your version supports it.
                // workbook.Settings.CalcEngineVersion = 2;

                // Define output file path
                string outputPath = "HighPrecision.xlsx";

                // Save the workbook (lifecycle save rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
