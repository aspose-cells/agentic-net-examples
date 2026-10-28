// Title: How to verify a PivotTable has at least one data field before refreshing and calculating with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells check PivotTable data fields before calling CalculateData | validate that a PivotTable contains data fields using Aspose.Cells .NET | how to ensure a PivotTable has at least one data field in Aspose.Cells C# | Aspose.Cells .NET prevent PivotTable calculation when no data fields are defined | sample code to test PivotTable data field count with Aspose.Cells
// Tags: Aspose.Cells PivotTable data field validation | C# PivotTable data field count check | Aspose.Cells RefreshData and CalculateData example | Excel PivotTable existence verification .NET | Aspose.Cells workbook save after PivotTable processing

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, confirms the file exists, ensures the first worksheet contains a PivotTable, validates that the PivotTable has at least one data field, then refreshes and recalculates the PivotTable before saving the workbook.
class PivotTableValidator
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Get the first PivotTable (or iterate as required)
            PivotTable pivot = sheet.PivotTables[0];

            // Validate that the PivotTable has at least one data field
            if (pivot.DataFields.Count == 0)
            {
                Console.WriteLine("The PivotTable does not contain any data fields. Add at least one before performing calculations.");
                return;
            }

            // PivotTable validation succeeded – proceed with calculations
            Console.WriteLine("PivotTable validation passed. Performing calculations...");

            // Refresh data and recalculate the PivotTable
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
