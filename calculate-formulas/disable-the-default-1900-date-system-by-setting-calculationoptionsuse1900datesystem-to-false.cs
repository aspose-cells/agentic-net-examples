// Title: How to disable the 1900 date system and enable the 1904 date system in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that sets CalculationOptions.Use1900DateSystem = false, recalculates all formulas, and saves the workbook with Aspose.Cells. | Show how to switch a workbook to the 1904 date system via workbook.Settings.Date1904 = true, then recalculate formulas and export to an XLSX file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# disable 1900 date system and enable 1904 | set CalculationOptions.Use1900DateSystem false example Aspose.Cells | switch workbook to 1904 date system and recalculate formulas Aspose.Cells .NET | save workbook after changing date system Aspose.Cells C# | how to change Excel date system programmatically with Aspose.Cells
// Tags: disable 1900 date system Aspose.Cells | enable 1904 date system workbook C# | set CalculationOptions.Use1900DateSystem false | recalculate formulas after date system change | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, disables the default 1900 date system by enabling the 1904 system (workbook.Settings.Date1904 = true), recalculates all formulas with the default calculation options, and saves the result to 'Workbook_Disable1900DateSystem.xlsx', handling any exceptions that may occur.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (default 1900 date system is enabled)
                Workbook workbook = new Workbook();

                // Switch to the 1904 date system (disables the 1900 date system)
                workbook.Settings.Date1904 = true;

                // Recalculate formulas using default calculation options
                workbook.CalculateFormula();

                // Save the workbook to a file
                string outputFile = "Workbook_Disable1900DateSystem.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
