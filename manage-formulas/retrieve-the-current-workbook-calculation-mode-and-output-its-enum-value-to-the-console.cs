// Title: Read and print the workbook's CalcMode enumeration using Aspose.Cells for .NET in C#
// Common Searches: how to get the CalcMode of an Excel workbook using Aspose.Cells in C# | retrieve the current calculation mode from a workbook with Aspose.Cells .NET | display workbook calculation setting in console using Aspose.Cells | read Excel calculation mode enum with Aspose.Cells API | C# code to obtain workbook calculation mode without opening the file in UI
// Tags: Aspose.Cells workbook calculation mode retrieval | C# Aspose.Cells CalcMode property access | Excel workbook settings enumeration Aspose.Cells | console output of calculation mode Aspose.Cells | reflection fallback for CalcMode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, obtains the CalcMode value from the workbook's Settings (using reflection for version compatibility), and writes the calculation mode enum (or a not‑available notice) to the console while handling potential exceptions.
    class Program
    {
        static void Main()
        {
            const string inputFile = "input.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the specified file
                var workbook = new Workbook(inputFile);

                // Attempt to retrieve the calculation mode via reflection (compatible with multiple Aspose.Cells versions)
                var settingsType = workbook.Settings.GetType();
                var calcModeProp = settingsType.GetProperty("CalcMode");
                var calcModeValue = calcModeProp != null ? calcModeProp.GetValue(workbook.Settings) : null;

                // Output the calculation mode (or indicate unavailability)
                Console.WriteLine($"Calculation mode: {calcModeValue ?? "Not available"}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
