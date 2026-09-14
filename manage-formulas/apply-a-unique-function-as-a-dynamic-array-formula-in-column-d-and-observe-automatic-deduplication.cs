// Title: Use Aspose.Cells for .NET to apply the UNIQUE dynamic array formula in column D for automatic deduplication of a range
// AI Prompts: Write C# code with Aspose.Cells that places the formula =UNIQUE(A1:C3) into cell D1, forces calculation, and saves the workbook. | Show how to trigger dynamic array spilling in Aspose.Cells by inserting the UNIQUE function and then programmatically read the deduplicated list.
// Common Searches: asp.net apply UNIQUE function to Excel range using Aspose.Cells | c# dynamic array formula UNIQUE spill results column D Aspose.Cells example | how to deduplicate cells A1:C3 with UNIQUE in Aspose.Cells | recalculate formulas after setting UNIQUE in an Aspose.Cells workbook
// Tags: Aspose.Cells set UNIQUE formula | C# dynamic array UNIQUE Excel | deduplicate range with Aspose.Cells | spill array results column D Aspose.Cells | calculate formulas Aspose.Cells workbook

using Aspose.Cells;
using System;

// The program creates a new workbook, fills A1:C3 with sample data containing duplicates, assigns the UNIQUE dynamic array formula to D1 to automatically deduplicate and spill the results, recalculates all formulas, and saves the file as UniqueDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];

        // Sample data in columns A, B, C (including duplicates)
        worksheet.Cells["A1"].PutValue("Apple");
        worksheet.Cells["B1"].PutValue("Banana");
        worksheet.Cells["C1"].PutValue("Apple");

        worksheet.Cells["A2"].PutValue("Orange");
        worksheet.Cells["B2"].PutValue("Banana");
        worksheet.Cells["C2"].PutValue("Grape");

        worksheet.Cells["A3"].PutValue("Apple");
        worksheet.Cells["B3"].PutValue("Lemon");
        worksheet.Cells["C3"].PutValue("Orange");

        // Apply the UNIQUE dynamic array formula in D1.
        // This will automatically deduplicate the values from A1:C3 and spill the results into column D.
        worksheet.Cells["D1"].Formula = "=UNIQUE(A1:C3)";

        // Recalculate formulas so the UNIQUE function is evaluated
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("UniqueDemo.xlsx");
    }
}
