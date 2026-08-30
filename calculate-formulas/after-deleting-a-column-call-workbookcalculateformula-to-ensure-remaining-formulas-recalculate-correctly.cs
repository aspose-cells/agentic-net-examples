// Title: Delete a column and recalculate dependent formulas in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to remove column B from a worksheet and automatically adjust any formulas that reference the removed column. | After removing a column, invoke the method that recomputes all workbook formulas and then save the updated .xlsx file. | Show how to shift cells left when deleting a column and ensure formulas like =A1+B1 update correctly with Aspose.Cells.
// Common Searches: Aspose.Cells C# delete a column and update formulas automatically | How to recalculate Excel formulas after column removal with Aspose.Cells .NET | Workbook.CalculateFormula usage after structural changes in Aspose.Cells | Shift cells left and refresh dependent formulas using Aspose.Cells C# | Delete column B and keep =A1+B1 correct in Aspose.Cells
// Tags: delete column Aspose.Cells recalculate formulas | recalculate workbook formulas after column removal C# | adjust formula references Aspose.Cells delete column | Workbook.CalculateFormula after structural change .NET | shift cells left delete column Aspose.Cells | update dependent formulas C# Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, writes values to A1 and B1, sets C1 to =A1+B1, deletes column B while shifting cells left, calls Workbook.CalculateFormula to recompute the formula, and saves the file as DeleteColumnRecalcDemo.xlsx.
    public class DeleteColumnAndRecalculateDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Add sample data and a formula that references columns A and B
                cells["A1"].PutValue(10);
                cells["B1"].PutValue(20);
                cells["C1"].Formula = "=A1+B1";

                // Delete column B (index 1) and update references in formulas
                cells.DeleteColumn(1, true);

                // Recalculate all formulas after the column deletion
                workbook.CalculateFormula();

                // Save the modified workbook
                string outputPath = "DeleteColumnRecalcDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeleteColumnAndRecalculateDemo.Run();
        }
    }
}
