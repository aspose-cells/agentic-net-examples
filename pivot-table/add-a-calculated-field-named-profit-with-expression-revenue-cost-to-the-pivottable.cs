// Title: Add a Profit calculated field ([Revenue]-[Cost]) to a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a calculated field named Profit with the expression [Revenue]-[Cost] into the first PivotTable of a workbook via Aspose.Cells C#. | Refresh the PivotTable after adding the calculated field and write the updated workbook to a new Excel file. | Check for the existence of the source file and handle exceptions when loading the workbook.
// Common Searches: Aspose.Cells C# add profit calculated field to existing pivot table | C# code to create calculated field [Revenue]-[Cost] in Excel pivot using Aspose | Refresh pivot table after adding calculated field Aspose.Cells .NET example | Handle missing input.xlsx when loading workbook with Aspose.Cells C#
// Tags: calculated field insertion pivot Aspose.Cells | profit formula [Revenue]-[Cost] C# | pivot table data refresh Aspose.Cells | export updated workbook Aspose.Cells .NET | input file validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads an existing Excel workbook, accesses its first worksheet and PivotTable, adds a calculated field called Profit defined as [Revenue]-[Cost], refreshes the PivotTable to apply the change, and saves the result to a new file while safely handling missing input files and other exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a calculated field named "Profit" with the expression [Revenue]-[Cost]
            pivotTable.AddCalculatedField("Profit", "[Revenue]-[Cost]");

            // Refresh the PivotTable to apply the new calculated field
            pivotTable.RefreshData();

            // Save the modified workbook to a new file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
