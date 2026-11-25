// Title: Set each worksheet's TabId to its position and save the workbook using Aspose.Cells for .NET
// Common Searches: how to assign TabId to worksheets based on their index with Aspose.Cells C# | Aspose.Cells C# set worksheet TabId sequentially and save workbook | update Excel sheet tab order programmatically using Aspose.Cells .NET | C# code to change worksheet TabId values in an existing workbook | save workbook after modifying TabId property with Aspose.Cells
// Tags: Aspose.Cells worksheet TabId assignment | C# sequential TabId for Excel sheets | modify Excel tab order with Aspose.Cells | save workbook after TabId update Aspose.Cells | set TabId property using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace SetWorksheetTabIds
{
    // The example loads an existing Excel file (or creates a new workbook), iterates through all worksheets, sets each sheet's TabId to its zero‑based index plus one to define tab order, and then saves the workbook, applying the new ordering.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = args.Length > 0 ? args[0] : string.Empty;
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            try
            {
                Workbook workbook;

                if (!string.IsNullOrWhiteSpace(inputPath))
                {
                    if (!File.Exists(inputPath))
                    {
                        Console.Error.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                        return;
                    }

                    Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine("No input file specified. Creating a new workbook with sample worksheets.");
                    workbook = new Workbook();

                    // Ensure there are at least three worksheets for demonstration
                    workbook.Worksheets[0].Name = "Sheet1";
                    workbook.Worksheets.Add("Sheet2");
                    workbook.Worksheets.Add("Sheet3");
                }

                // Iterate through all worksheets and set TabId = Index + 1
                Console.WriteLine("Setting TabId for each worksheet...");
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    int newTabId = i + 1; // Index is zero‑based, TabId expects a positive integer
                    sheet.TabId = newTabId;
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" (Index={sheet.Index}) => TabId set to {newTabId}");
                }

                // Save the workbook to apply the changes
                Console.WriteLine($"Saving workbook to \"{outputPath}\"...");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Save completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
