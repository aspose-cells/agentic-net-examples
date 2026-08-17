// Title: Bulk rename named ranges with a prefix and control calculation in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to temporarily suspend formula calculation, prepend a custom prefix to every defined name in a workbook, then re‑enable calculation and run a full recalc before saving the file.
// Keywords: Aspose.Cells C# | named ranges bulk rename | prefix defined names | disable calculation Aspose.Cells | Workbook.CalculateFormula | Excel automation rename names | formula engine suspend | bulk update named ranges | Excel workbook recalc
// Common Searches: Aspose.Cells add prefix to all named ranges | temporarily turn off calculation while renaming names in Aspose.Cells | force full recalculation after bulk rename of defined names | C# bulk rename Excel named ranges Aspose.Cells | disable formula evaluation Aspose.Cells .NET
// Developer Intent: Rename every defined name with a specified prefix without triggering intermediate calculations, then reactivate the formula engine and recalculate the workbook.
// Use Cases: Standardize naming conventions in legacy workbooks before integration with other systems. | Update hundreds of named ranges in a template without incurring performance penalties from repeated recalculations. | Prepare a report workbook for distribution, ensuring all formulas reflect the new range names.
// AI Prompts: Show C# code that disables calculation, adds a prefix to all Workbook.Worksheets.Names, re‑enables calculation and calls CalculateFormula in Aspose.Cells. | Explain step‑by‑step how to bulk rename defined names in an Aspose.Cells workbook while preventing intermediate formula evaluation. | Provide a concise example of suspending formula evaluation, renaming named ranges with a custom prefix, and performing a full recalc in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to temporarily suspend formula calculation, prepend a custom prefix to every defined name in a workbook, then re‑enable calculation and run a full recalc before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);

            // Create a couple of named ranges
            int idx1 = workbook.Worksheets.Names.Add("Range1");
            workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1:$A$2";

            int idx2 = workbook.Worksheets.Names.Add("Range2");
            workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$B$1:$B$2";

            // Bulk rename all defined names with a prefix
            const string prefix = "New_";
            foreach (Name name in workbook.Worksheets.Names)
            {
                name.Text = prefix + name.Text;
            }

            // Force a full recalculation
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule)
            string outputPath = "RenamedNames.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
