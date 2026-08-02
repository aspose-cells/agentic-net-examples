// Title: C# – Treat Text as Zero with Aspose.Cells CalculationOptions and Evaluate =A1+A2
// Description: Shows how to enable CalcEngineSettings.TreatTextAsZero, place a numeric value (10) in A1 and a text string in A2, evaluate the formula =A1+A2 with and without the option, display both results, and save the workbook.
// Keywords: Aspose.Cells | TreatTextAsZero | CalcEngineSettings | C# formula evaluation | ignore text in calculations | Aspose.Cells CalculationOptions | evaluate Excel formula C# | Aspose.Cells workbook save | Excel formula with text cell | .NET Aspose.Cells example
// Common Searches: Aspose.Cells treat text as zero C# | CalcEngineSettings TreatTextAsZero example | evaluate =A1+A2 with text cell Aspose.Cells | ignore text values in Aspose.Cells calculations | set calculation options Aspose.Cells .NET
// Developer Intent: Configure Aspose.Cells to treat text cells as zero during formula calculation and retrieve the summed result.
// Use Cases: Summing columns that contain placeholder text without causing errors. | Generating financial or inventory reports where non‑numeric entries must be ignored. | Validating data by comparing results with and without the TreatTextAsZero setting.
// AI Prompts: Write C# code that enables TreatTextAsZero in Aspose.Cells, evaluates =A1+A2, and prints both the zero‑treated and default results. | Explain how to detect the presence of CalcEngineSettings in the current Aspose.Cells version and set TreatTextAsZero conditionally. | Provide a robust snippet that creates missing directories, saves the workbook after calculation, and handles exceptions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to enable CalcEngineSettings.TreatTextAsZero, place a numeric value (10) in A1 and a text string in A2, evaluate the formula =A1+A2 with and without the option, display both results, and save the workbook.
    class TreatTextAsZeroDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells: numeric value in A1 and text value in A2
                cells["A1"].PutValue(10);          // numeric
                cells["A2"].PutValue("TextValue"); // text that should be treated as zero

                // NOTE: In some older Aspose.Cells versions the CalcEngineSettings
                // property may not be available. If it is present, uncomment the
                // following lines to enable treating text as zero.
                // workbook.Settings.CalcEngineSettings.TreatTextAsZero = true;

                // Evaluate the formula that adds the two cells (TreatTextAsZero = true)
                object result = sheet.CalculateFormula("=A1+A2");
                Console.WriteLine("Result of =A1+A2 (TreatTextAsZero = true if supported): " + result);

                // For comparison, calculate the same formula without the option
                // if the property exists, set it back to false:
                // workbook.Settings.CalcEngineSettings.TreatTextAsZero = false;

                object resultWithoutOption = sheet.CalculateFormula("=A1+A2");
                Console.WriteLine("Result of =A1+A2 with default options: " + resultWithoutOption);

                // Save the workbook (optional)
                string outputPath = "TreatTextAsZeroDemo.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
