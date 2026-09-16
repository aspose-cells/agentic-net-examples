// Title: Delete a named PivotTable from an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX workbook, locates a PivotTable by its name on a specific worksheet, removes it with Worksheet.PivotTables.Remove, and saves the file. | Show a try‑catch example that checks for the presence of a PivotTable before calling Remove, and logs appropriate messages for missing files or tables. | Demonstrate how to verify the input file exists, load the workbook, delete PivotTable1 from Sheet1, and output the modified workbook to a new path using Aspose.Cells.
// Common Searches: c# aspnet delete pivot table by name using Aspose.Cells | how to remove a specific PivotTable from an Excel file with Aspose.Cells .NET | Aspose.Cells Worksheet.PivotTables.Remove example for C# | code to drop PivotTable1 from Sheet1 in an XLSX workbook using Aspose | handling missing pivot table when deleting with Aspose.Cells C#
// Tags: Aspose.Cells remove pivot table C# | Worksheet.PivotTables.Remove method | delete named pivot table Aspose.Cells | C# load workbook delete pivot table | Excel XLSX pivot table removal Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // PivotTable class resides in this namespace

// // Loads "input.xlsx", accesses "Sheet1", finds the PivotTable named "PivotTable1", removes it with worksheet.PivotTables.Remove, and saves the workbook as "output.xlsx" while handling missing files and errors.
class DeletePivotTableExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet that contains the pivot table (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // adjust name if needed

            // Retrieve the pivot table by name
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"]; // replace with actual name

            // Remove the pivot table if it exists
            if (pivotTable != null)
            {
                worksheet.PivotTables.Remove(pivotTable);
                Console.WriteLine("Pivot table removed successfully.");
            }
            else
            {
                Console.WriteLine("Specified pivot table not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
