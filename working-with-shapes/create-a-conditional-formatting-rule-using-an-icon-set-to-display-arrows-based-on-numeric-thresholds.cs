// Title: Apply a 3‑arrow Icon Set conditional formatting with custom numeric thresholds to a range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that applies an arrow‑based icon set to the range A1:A10 and defines thresholds at 0, 5, and 10. | Demonstrate how to set IconSet options like type, direction, value display, and individual criteria formulas with Aspose.Cells.
// Common Searches: how to add an arrow icon set conditional formatting in Aspose.Cells C# | Aspose.Cells set custom numeric thresholds for IconSetCondition | C# example for applying three‑arrow icon set to a column in Excel with Aspose.Cells | configure IconSetCondition thresholds using formulas in Aspose.Cells .NET
// Tags: Aspose.Cells conditional formatting icon set | C# apply three‑arrow icon set | numeric thresholds for icon set in Aspose.Cells | Excel conditional formatting arrows .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills cells A1:A10 with even numbers, defines that range for conditional formatting, adds an IconSet rule of type Arrows3, configures it to show values, sets the direction, and establishes three numeric criteria (>=10, >=5, >=0) before saving the file as ConditionalFormatting_IconSet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 2); // 0,2,4,...,18
            }

            // Define the range for conditional formatting (A1:A10)
            CellArea range = CellArea.CreateCellArea("A1", "A10");

            // Add a new conditional formatting rule
            int cfIndex = sheet.ConditionalFormattings.Add();
            // Use dynamic to avoid compile‑time type issues with ConditionalFormatting
            dynamic cf = sheet.ConditionalFormattings[cfIndex];
            cf.Area = range;

            // Add an Icon Set condition (arrows)
            cf.AddCondition(FormatConditionType.IconSet);
            // Use dynamic for IconSetCondition as well
            dynamic iconSet = cf.IconSet;

            // Configure the 3‑arrow icon set
            iconSet.IconSetType = IconSetType.Arrows3;
            iconSet.Reverse = false;      // higher values show upward arrows
            iconSet.ShowValue = true;     // display cell value with the icon

            // Define thresholds for the three icons
            iconSet.IconCriteria[0].Operator = OperatorType.GreaterOrEqual;
            iconSet.IconCriteria[0].Formula = "10";

            iconSet.IconCriteria[1].Operator = OperatorType.GreaterOrEqual;
            iconSet.IconCriteria[1].Formula = "5";

            iconSet.IconCriteria[2].Operator = OperatorType.GreaterOrEqual;
            iconSet.IconCriteria[2].Formula = "0";

            // Save the workbook
            string outputPath = "ConditionalFormatting_IconSet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
