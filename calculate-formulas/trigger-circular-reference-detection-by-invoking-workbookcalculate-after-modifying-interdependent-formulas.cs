// Title: How to detect circular references in Excel formulas using Workbook.CalculateFormula with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# detect circular reference during workbook calculation | trigger Workbook.CalculateFormula to catch circular reference errors in .NET | example of handling interdependent formulas that cause circular references in Aspose.Cells | use Workbook.CalculateFormula to validate Excel formulas for circular loops in C# | catch circular reference exception when saving a workbook with Aspose.Cells
// Tags: circular-reference detection Aspose.Cells | Workbook.CalculateFormula usage | interdependent formulas Excel Aspose.Cells | exception handling workbook calculation | saving workbook after formula evaluation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

namespace Demo
{
    // The sample creates a workbook, sets interdependent formulas in cells A1 and B1 that form a circular reference, calls Workbook.CalculateFormula to trigger calculation and detect the loop, catches the resulting exception, and then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set interdependent formulas that create a circular reference
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            try
            {
                // Trigger calculation; this will detect the circular reference
                workbook.CalculateFormula();
                Console.WriteLine("Calculation completed without circular reference.");
            }
            catch (Exception ex)
            {
                // Handle any calculation errors, including circular references
                Console.WriteLine("Error during calculation: " + ex.Message);
            }

            try
            {
                // Save the workbook (optional)
                string outputPath = "CircularReferenceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
    }
}
