// Title: Print gridlines on the first three worksheets with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds extra sheets, and enables PageSetup.PrintGridlines only on the first three worksheets before saving the file as GridlinesFirstThreeSheets.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PrintGridlines | gridlines printing | first three worksheets | page setup | conditional worksheet settings | workbook sheet index | Excel gridline control
// Common Searches: Aspose.Cells enable gridlines on selected sheets | C# set PrintGridlines for first three worksheets | How to print gridlines only on certain worksheets using Aspose.Cells | Conditional gridline printing Aspose.Cells .NET | Hide gridlines on some Excel sheets with Aspose.Cells
// Developer Intent: Apply the PrintGridlines flag to the first three worksheets while leaving all other sheets unchanged.
// Use Cases: Generate a multi‑sheet report where only the cover and summary pages need printed gridlines. | Create an invoice workbook where the first three pages require gridlines for alignment, but detailed data sheets do not. | Prepare a template that mixes formatted and raw data sheets, enabling gridlines only on the setup worksheets.
// AI Prompts: Write C# code using Aspose.Cells that sets PageSetup.PrintGridlines = true for the first N worksheets of a workbook. | Show how to toggle the PrintGridlines property based on worksheet index or name in Aspose.Cells for .NET. | Explain a method to conditionally enable gridline printing for a subset of worksheets while preserving default settings for the rest.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds extra sheets, and enables PageSetup.PrintGridlines only on the first three worksheets before saving the file as GridlinesFirstThreeSheets.xlsx.
    class EnableGridlinesFirstThreeSheets
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets so we have at least five sheets
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4");
            workbook.Worksheets.Add("Sheet5");

            // Enable printing of gridlines for the first three worksheets only
            int sheetsToEnable = Math.Min(3, workbook.Worksheets.Count);
            for (int i = 0; i < sheetsToEnable; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                ws.PageSetup.PrintGridlines = true; // Gridlines will be printed when this sheet is printed
            }

            // Save the workbook to a file
            string outputPath = "GridlinesFirstThreeSheets.xlsx";
            workbook.Save(outputPath);
        }
    }
}
