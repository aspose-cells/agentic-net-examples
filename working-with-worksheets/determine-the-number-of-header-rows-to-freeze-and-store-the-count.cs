// Title: How to read the number of frozen header rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, accesses the first worksheet, and returns the count of frozen rows, using reflection to handle library versions where FreezePanes may be a property. | Provide a .NET snippet that checks for the existence of Worksheet.FreezePanes and extracts its RowCount to determine how many header rows are frozen.
// Common Searches: C# Aspose.Cells get frozen rows count from worksheet | how to determine number of frozen header rows in Excel using Aspose.Cells .NET | read FreezePanes.RowCount with Aspose.Cells for .NET | Aspose.Cells retrieve frozen pane rows via reflection | check if FreezePanes property exists Aspose.Cells C#
// Tags: aspocells read freeze panes row count | c# get frozen header rows aspocells | worksheet freezepanes reflection aspocells | excel workbook frozen rows detection .net | aspocells compatibility read freezepanes property

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, accesses the first worksheet, and uses reflection to safely obtain the FreezePanes.RowCount (when available) to determine how many header rows are frozen, storing the result in a variable and printing it.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            int headerRowsToFreeze = 0;

            // Attempt to read the number of frozen rows using reflection.
            // This works for versions where Worksheet.FreezePanes is a property.
            try
            {
                var freezeProp = typeof(Worksheet).GetProperty("FreezePanes");
                if (freezeProp != null)
                {
                    var freezeObj = freezeProp.GetValue(sheet);
                    var rowCountProp = freezeObj?.GetType().GetProperty("RowCount");
                    if (rowCountProp != null)
                    {
                        headerRowsToFreeze = (int)rowCountProp.GetValue(freezeObj);
                    }
                }
            }
            catch
            {
                // Property not available or other issue; keep default value (0)
                headerRowsToFreeze = 0;
            }

            Console.WriteLine($"Number of header rows frozen: {headerRowsToFreeze}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
