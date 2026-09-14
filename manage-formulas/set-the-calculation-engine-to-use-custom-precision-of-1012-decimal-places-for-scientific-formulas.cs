// Title: How to set custom decimal precision (10‑12 places) for scientific formulas using Aspose.Cells CalcEngineSettings in C#
// AI Prompts: Configure the Aspose.Cells calculation engine to disable PrecisionAsDisplayed and set NumberDecimalPlaces to 12 via CalcEngineSettings in C#. | Use reflection to access CalcEngineSettings when it is not directly exposed, then apply a custom decimal precision to a workbook. | Add robust error handling for missing input files and unavailable CalcEngineSettings while adjusting formula precision.
// Common Searches: Aspose.Cells C# set number of decimal places for formula calculation | How to change calculation precision for scientific formulas in an Aspose.Cells workbook | Disable PrecisionAsDisplayed and set custom decimal places with Aspose.Cells .NET | Reflection example to modify CalcEngineSettings properties in Aspose.Cells | Error handling for missing CalcEngineSettings API in Aspose.Cells
// Tags: calcenginesettings decimal precision Aspose.Cells | disable precisionasdisplayed property .NET | set numberdecimalplaces via reflection | custom formula precision Excel workbook C# | handle missing CalcEngineSettings API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, uses reflection to obtain the CalcEngineSettings object, disables the PrecisionAsDisplayed flag, sets the NumberDecimalPlaces to 12 for higher scientific formula accuracy, and saves the workbook. It also includes checks for file existence and graceful handling when the CalcEngineSettings API is unavailable.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Attempt to configure custom precision if the API is available in the referenced version
                var calcEngineProp = workbook.Settings.GetType().GetProperty("CalcEngineSettings");
                if (calcEngineProp != null)
                {
                    object calcEngine = calcEngineProp.GetValue(workbook.Settings);
                    var precisionProp = calcEngine?.GetType().GetProperty("PrecisionAsDisplayed");
                    var decimalPlacesProp = calcEngine?.GetType().GetProperty("NumberDecimalPlaces");

                    if (precisionProp != null)
                    {
                        precisionProp.SetValue(calcEngine, false);
                    }

                    if (decimalPlacesProp != null)
                    {
                        decimalPlacesProp.SetValue(calcEngine, 12);
                    }
                }

                // Save the workbook to the specified output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
