// Title: Retrieve Worksheet TabId (SheetId) for Every Sheet in an Aspose.Cells .NET Workbook
// Description: Creates an in‑memory workbook, adds several worksheets, then loops through the Worksheets collection to read each sheet's internal TabId and name, outputs the values to the console, and saves the file.
// Keywords: Aspose.Cells TabId | worksheet SheetId .NET | iterate worksheets Aspose | get internal worksheet identifier | Aspose.Cells diagnostic IDs
// Common Searches: Aspose.Cells get worksheet TabId C# | how to read SheetId of each worksheet | list all worksheet IDs Aspose.Cells | C# iterate worksheets and show TabId
// Developer Intent: Extract the internal TabId of every worksheet in a workbook for logging or validation.
// Use Cases: Log TabIds after adding or removing sheets to confirm unique identifiers. | Build a name‑to‑ID map for custom metadata or external references. | Detect changes in worksheet ordering by comparing TabIds before and after modifications.
// AI Prompts: Generate a C# function that returns a Dictionary<string, int> mapping worksheet names to their TabIds using Aspose.Cells. | Write code that iterates all worksheets in a workbook and writes each Name and TabId to a CSV file. | Explain the difference between a worksheet's TabId and its index in Aspose.Cells and when each should be used.

using System;
using Aspose.Cells;

namespace WorksheetIdDiagnostic
{
    // Creates an in‑memory workbook, adds several worksheets, then loops through the Worksheets collection to read each sheet's internal TabId and name, outputs the values to the console, and saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Rename the default first worksheet
                workbook.Worksheets[0].Name = "FirstSheet";

                // Add additional worksheets and keep references
                Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
                Worksheet thirdSheet = workbook.Worksheets.Add("ThirdSheet");

                // Optionally put some data (not required for ID retrieval)
                secondSheet.Cells["A1"].PutValue("Data in second sheet");
                thirdSheet.Cells["A1"].PutValue("Data in third sheet");

                // Iterate over the worksheet collection and obtain each TabId (internal sheet identifier)
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int sheetId = sheet.TabId;               // Internal identifier for the sheet
                    string sheetName = sheet.Name;           // Friendly name of the sheet
                    Console.WriteLine($"Worksheet \"{sheetName}\" has TabId: {sheetId}");
                }

                // Save the workbook to verify that everything works (optional)
                string outputPath = "WorksheetIdDiagnostic.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
