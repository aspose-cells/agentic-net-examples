// Title: Add red font conditional formatting for cells containing the word “Error” in a specific range using Aspose.Cells for .NET (C#)
// AI Prompts: Create a ContainsText conditional formatting rule that changes the font color to red for any cell containing the word 'Error' within A1:Z100. | Add a conditional formatting collection to a worksheet and set a red font style for cells whose text includes 'Error' using Aspose.Cells in C#. | Modify an existing workbook so that all cells with the word 'Error' are highlighted with red font via Aspose.Cells conditional formatting.
// Common Searches: asp.net aspose.cells conditional formatting contains text red font | how to highlight cells with the word error using Aspose.Cells C# | apply red font conditional format to Excel range A1 Z100 Aspose.Cells | Aspose.Cells FormatConditionType.ContainsText example for error highlighting | C# code to add conditional formatting for error messages in Excel workbook
// Tags: Aspose.Cells conditional formatting ContainsText | red font style for error text Aspose.Cells | apply conditional formatting to Excel range A1-Z100 C# | highlight error cells using Aspose.Cells .NET | FormatCondition Type ContainsText example

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program loads or creates a workbook, defines a conditional formatting rule for the range A1:Z100 that sets the font color to red when a cell contains the word "Error", and saves the updated file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range where the rule should be applied
            CellArea area = CellArea.CreateCellArea("A1", "Z100");

            // Add a new conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];
            fcc.AddArea(area);

            // Add a condition that checks if the cell contains the word "Error"
            // OperatorType.None and an empty second formula are required by the API
            int conditionIndex = fcc.AddCondition(FormatConditionType.ContainsText, OperatorType.None, "Error", string.Empty);

            // Set the font color of cells meeting the condition to red
            FormatCondition condition = fcc[conditionIndex];
            condition.Style.Font.Color = Color.Red;

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
