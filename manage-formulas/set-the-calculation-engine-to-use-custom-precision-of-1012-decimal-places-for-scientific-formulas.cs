// Title: Configure Aspose.Cells to use Decimal precision for 10‑12 decimal places (C#)
// Description: Demonstrates how to set Aspose.Cells' calculation engine to the Decimal precision strategy, enabling 10‑12 decimal place accuracy for scientific formulas. The example creates a workbook, inserts a high‑precision formula, applies CalculationOptions.PrecisionStrategy = Decimal, calculates the result, and saves the file.
// Keywords: Aspose.Cells Decimal precision | CalculationOptions PrecisionStrategy | high precision scientific formula C# | 10‑12 decimal places Aspose.Cells | custom calculation accuracy .NET | Aspose.Cells calculation engine | decimal vs double precision Aspose
// Common Searches: Aspose.Cells set decimal precision for formulas | CalculationPrecisionStrategy.Decimal example | increase formula accuracy Aspose.Cells .NET | how to calculate small scientific constants with Aspose.Cells | custom precision options Aspose.Cells C#
// Developer Intent: Enable decimal‑based calculation in Aspose.Cells to achieve 10‑12 decimal place accuracy for scientific formulas.
// Use Cases: Calculating tiny scientific constants (e.g., π × 10⁻¹¹) without double‑precision rounding errors. | Performing engineering or financial analyses that require consistent high‑precision results across worksheets. | Generating reports where regulatory standards demand exact decimal representation of computed values.
// AI Prompts: Show C# code that configures Aspose.Cells CalculationOptions to use Decimal precision and evaluates a workbook. | Explain how CalculationPrecisionStrategy.Decimal provides up to 28‑29 significant digits compared to Double. | Give a step‑by‑step guide to verify 10‑12 decimal place results after setting custom precision in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells' calculation engine to the Decimal precision strategy, enabling 10‑12 decimal place accuracy for scientific formulas. The example creates a workbook, inserts a high‑precision formula, applies CalculationOptions.PrecisionStrategy = Decimal, calculates the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example of a scientific formula that requires high precision
        // This calculates 3.1415926535 × 10⁻¹¹
        worksheet.Cells["A1"].Formula = "=POWER(10, -11) * 3.1415926535";

        // Configure calculation options to use decimal precision.
        // The Decimal strategy processes operands as System.Decimal,
        // providing up to 28‑29 significant digits (covers the 10‑12 decimal place requirement).
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Perform the calculation with the custom precision settings
        workbook.CalculateFormula(calcOptions);

        // Output the calculated value
        Console.WriteLine("Calculated result (A1): " + worksheet.Cells["A1"].Value);

        // Save the workbook (lifecycle rule)
        workbook.Save("CustomPrecision.xlsx");
    }
}
