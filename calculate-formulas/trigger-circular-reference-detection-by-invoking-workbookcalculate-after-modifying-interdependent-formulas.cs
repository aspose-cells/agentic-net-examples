// Title: Trigger circular reference detection in Excel formulas using Aspose.Cells Workbook.CalculateFormula (C#)
// Common Searches: how to catch circular reference exception when calculating formulas with Aspose.Cells in C# | using Workbook.CalculateFormula to identify interdependent formulas in Aspose.Cells | Aspose.Cells example for detecting circular references in an Excel workbook | trigger formula calculation to find circular references in .NET Aspose.Cells | detect circular reference error after setting A1 and B1 formulas with Aspose.Cells
// Tags: Workbook.CalculateFormula API for circular reference handling | Aspose.Cells interdependent formula example | C# Excel circular reference exception handling | Aspose.Cells save workbook to XLSX after calculation | Excel circular reference detection in .NET

using System;
using Aspose.Cells;

namespace CircularReferenceDemo
{
    // The sample creates a new workbook, assigns interdependent formulas (A1 = B1+1 and B1 = A1+1) that form a circular reference, and then calls Workbook.CalculateFormula inside a try‑catch block. Aspose.Cells throws an exception when the circular reference is detected, which is caught and reported. Finally, the workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set interdependent formulas that create a circular reference
                // A1 = B1 + 1
                sheet.Cells["A1"].Formula = "=B1+1";
                // B1 = A1 + 1
                sheet.Cells["B1"].Formula = "=A1+1";

                // Trigger calculation to detect the circular reference
                try
                {
                    // Correct API to calculate formulas
                    workbook.CalculateFormula();
                    Console.WriteLine("Calculation completed without detecting a circular reference.");
                }
                catch (Exception ex)
                {
                    // Aspose.Cells throws an exception for circular references
                    Console.WriteLine("Circular reference detected: " + ex.Message);
                }

                // Save the workbook
                workbook.Save("CircularReferenceDemo.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // General exception handling for unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
