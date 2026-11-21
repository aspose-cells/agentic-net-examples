// Title: Programmatically set the first PivotTable data field caption to "Sales Amount" using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change pivot table data field name to custom caption | set pivot table data field caption programmatically in .NET | rename first data field of an Excel pivot table using Aspose.Cells | update pivot table data caption and refresh pivot in C# | how to assign a custom caption to a pivot table data field with Aspose.Cells
// Tags: Aspose.Cells set pivot table data field caption | C# rename pivot table data field | refresh pivot table after caption change Aspose.Cells | calculate pivot table data Aspose.Cells C# | modify pivot table field name .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, accesses the first worksheet's PivotTable, changes the name of its first data field to "Sales Amount", then refreshes and recalculates the PivotTable before saving the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                // Retrieve the first PivotTable
                PivotTable pivotTable = sheet.PivotTables[0];

                // Set a caption for the first data field, if any
                if (pivotTable.DataFields.Count > 0)
                {
                    pivotTable.DataFields[0].Name = "Sales Amount";
                }

                // Refresh the PivotTable data (correct API)
                pivotTable.RefreshData();

                // Recalculate the PivotTable to apply changes
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("No PivotTables found in the first worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
