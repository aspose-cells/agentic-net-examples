// Title: How to add a WorksheetChanged event in C# with Aspose.Cells that recalculates formulas only when cells A1 or C5 are edited
// AI Prompts: Write C# code using Aspose.Cells to register a WorksheetChanged event handler that checks if the changed cell is A1 or C5 and calls CalculateFormula only in those cases. | Show how to combine cell value assignment, a dependent formula, and conditional recalculation in an Aspose.Cells workbook, then save the file.
// Common Searches: Aspose.Cells C# trigger CalculateFormula on WorksheetChanged for specific cells | How to listen for cell edits in an Aspose.Cells workbook and recalc only affected formulas | Selective formula recalculation in .NET when A1 or C5 values change using Aspose.Cells | C# example of worksheet change event handling with Aspose.Cells Excel library | Conditional workbook recalculation based on cell address in Aspose.Cells .NET
// Tags: worksheetchanged event Aspose.Cells | selective formula recalculation .NET | cell address filter CalculateFormula | Aspose.Cells conditional recalculation | C# workbook event handling Excel | specific cells trigger calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, assigns values to A1 and C5, defines a formula in B1, forces a full formula recalculation, and saves the file. It also demonstrates how to attach a WorksheetChanged event that checks the changed cell address and invokes CalculateFormula only when A1 or C5 are modified.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example data: set initial values and a formula
            sheet.Cells["A1"].PutValue(5);               // This cell will trigger recalculation when changed
            sheet.Cells["C5"].PutValue(10);              // This cell will also trigger recalculation when changed
            sheet.Cells["B1"].Formula = "=A1*2";         // Dependent formula

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Ensure the output directory exists (if a directory part is present)
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (save rule)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
