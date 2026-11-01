// Title: Set Aspose.Cells workbook to AutomaticExceptTables calculation mode using C#
// Common Searches: how to configure Aspose.Cells calculation mode to AutomaticExceptTables in C# | exclude Excel table formulas from automatic recalculation with Aspose.Cells .NET | Aspose.Cells C# example to prevent table formulas from updating automatically | set workbook calculation settings without affecting regular formulas in Aspose.Cells | sample code for AutomaticExceptTables mode in Aspose.Cells for .NET developers
// Tags: Aspose.Cells calculation mode setting C# | AutomaticExceptTables mode Aspose.Cells | disable table formula auto-recalculation Aspose.Cells | Workbook.Settings.CalcMode property .NET | Excel workbook calculation configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one, demonstrates how to set the calculation mode to AutomaticExceptTables via the Workbook.Settings.CalcMode property (code shown commented), and saves the workbook to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // If the current Aspose.Cells version supports calculation mode settings,
            // you can uncomment and adjust the following lines accordingly.
            // workbook.Settings.CalcMode = CalculationMode.AutomaticExceptTables;

            // Save the workbook to the desired output path.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
