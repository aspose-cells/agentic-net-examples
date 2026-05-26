using System;
using Aspose.Cells;

namespace RenameNamedRangesDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Iterate through the collection and rename those containing "Old"
            foreach (Name name in names)
            {
                if (name.Text != null && name.Text.Contains("Old"))
                {
                    // Create the new name by replacing "Old" with "New"
                    string newName = name.Text.Replace("Old", "New");

                    // Assign the new name back to the Name object
                    name.Text = newName;
                }
            }

            // Recalculate all formulas so that any dependent cells are updated
            workbook.CalculateFormula();

            // Save the modified workbook (replace with your desired output path)
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}