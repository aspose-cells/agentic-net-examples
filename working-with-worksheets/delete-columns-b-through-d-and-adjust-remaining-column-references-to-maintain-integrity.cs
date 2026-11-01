// Title: Delete columns B through D and shift remaining columns left in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Remove columns B, C, and D from the first worksheet and automatically shift the subsequent columns left with Aspose.Cells. | Invoke the DeleteColumns method with updateReferences set to true to delete a range of columns and keep formulas consistent in a C# project. | Create or load a workbook, delete a specific column range, and save the modified file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# delete columns B to D and shift left example | How to delete multiple adjacent columns and preserve formulas with Aspose.Cells .NET | DeleteColumns method updateReferences parameter usage in Aspose.Cells | Remove column range from Excel file using Aspose.Cells for .NET | C# code to delete columns and adjust column references in an .xlsx workbook
// Tags: Aspose.Cells DeleteColumns method usage | delete column range in .xlsx with C# | shift remaining columns after deletion Aspose.Cells | preserve formulas when removing columns .NET | C# Excel column removal example | update column references Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a sample workbook if it doesn't exist, loads it, deletes columns B through D on the first worksheet using DeleteColumns with updateReferences=true to shift remaining columns left and keep formulas intact, then saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple one if missing.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempSheet = tempWb.Worksheets[0];
                tempSheet.Cells["A1"].PutValue("A");
                tempSheet.Cells["B1"].PutValue("B");
                tempSheet.Cells["C1"].PutValue("C");
                tempSheet.Cells["D1"].PutValue("D");
                tempWb.Save(inputPath);
            }

            // Load the existing workbook.
            var workbook = new Workbook(inputPath);

            // Access the first worksheet.
            var sheet = workbook.Worksheets[0];

            // Delete columns B through D (indices 1,2,3). Update references.
            sheet.Cells.DeleteColumns(1, 3, true);

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
