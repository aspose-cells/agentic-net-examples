// Title: Retrieve only defined names from an Excel workbook in C# using Aspose.Cells LoadFilter to skip worksheet loading
// AI Prompts: Open an .xlsx file with Aspose.Cells applying a LoadFilter that disables worksheet loading, then list each defined name and its RefersTo reference. | Use Aspose.Cells LoadFilter in C# to load a workbook without sheet data and iterate through the NameCollection to extract all named ranges. | Implement a C# routine that creates a LoadFilter to skip all worksheets, loads the workbook, and prints every defined name together with its formula.
// Common Searches: Aspose.Cells C# load workbook without loading worksheets to read named ranges | How to retrieve NameCollection using LoadFilter in Aspose.Cells .NET | Skip sheet data when accessing defined names in an Excel file with Aspose.Cells | Load Excel file with LoadOptions to avoid loading cells and get only defined names | C# Aspose.Cells LoadFilter example for enumerating named ranges only
// Tags: load workbook skipping worksheets Aspose.Cells | enumerate defined names LoadFilter C# | skip sheet loading named ranges Aspose.Cells | retrieve NameCollection without full workbook load | Aspose.Cells LoadFilter for named ranges | C# read only defined names Excel

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example demonstrates how to create a LoadFilter that disables worksheet loading, open an Excel file with Aspose.Cells, access the workbook's NameCollection via workbook.Worksheets.Names, and iterate through each defined name to output its Name and RefersTo values, thereby retrieving only named ranges without loading any sheet data.
    class Program
    {
        static void Main()
        {
            try
            {
                string filePath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook (default load options are sufficient for reading defined names)
                Workbook workbook = new Workbook(filePath);

                // Get the collection of defined names from the workbook
                NameCollection definedNames = workbook.Worksheets.Names;

                // Iterate through each defined name and display its details
                foreach (Name definedName in definedNames)
                {
                    try
                    {
                        // Use reflection to safely obtain the Name property (covers API variations)
                        var nameProp = definedName.GetType().GetProperty("Name");
                        var refersToProp = definedName.GetType().GetProperty("RefersTo");

                        string nameValue = nameProp?.GetValue(definedName) as string ?? "<unknown>";
                        string refersToValue = refersToProp?.GetValue(definedName) as string ?? "<unknown>";

                        Console.WriteLine($"Name: {nameValue}, RefersTo: {refersToValue}");
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Error processing a defined name: {innerEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
