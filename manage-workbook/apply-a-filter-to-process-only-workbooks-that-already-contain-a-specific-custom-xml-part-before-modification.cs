// Title: Conditionally Process Workbooks with a Specific Custom XML Part Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, scans its CustomXmlPartCollection for a predefined GUID, and modifies the workbook (adds a worksheet) only when the required custom XML part is present, otherwise leaves the file unchanged.
// Keywords: Aspose.Cells | C# | custom XML part | conditional workbook processing | check XML part ID | add worksheet | Excel automation | GUID validation
// Common Searches: Aspose.Cells check custom XML part ID | C# filter Excel workbook by custom XML | process workbook only if XML part exists | conditional worksheet addition Aspose.Cells | validate custom XML part GUID in Excel
// Developer Intent: Detect a workbook that contains a particular custom XML part and apply changes only to those files.
// Use Cases: Validate incoming Excel templates that must embed a predefined custom XML schema before any business logic runs. | Skip files lacking required metadata to avoid errors in batch processing pipelines. | Append a summary sheet to reports that already carry a specific custom XML identifier. | Integrate with document‑management systems that tag spreadsheets with custom XML GUIDs.
// AI Prompts: Generate C# code with Aspose.Cells that reads a workbook, searches its CustomXmlParts for a given GUID, and adds a new worksheet only when the part is found. | Show how to log all custom XML part IDs in an Excel file before deciding whether to modify it using Aspose.Cells. | Write a reusable function that accepts a file path and a custom XML part GUID, returns true if the part exists, and conditionally inserts a worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlFilterDemo
{
    // Loads an Excel file, scans its CustomXmlPartCollection for a predefined GUID, and modifies the workbook (adds a worksheet) only when the required custom XML part is present, otherwise leaves the file unchanged.
    class Program
    {
        // ID of the custom XML part that must exist for the workbook to be processed
        private const string RequiredCustomXmlPartId = "2F087CB2-7CA8-43DA-B048-2E2F61F4936F";

        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";
            // Path where the modified workbook will be saved
            string outputPath = "output.xlsx";

            // Load the workbook (using the standard constructor as per lifecycle rules)
            Workbook workbook = new Workbook(sourcePath);

            // Check if the workbook contains the required custom XML part
            bool containsRequiredPart = false;
            CustomXmlPartCollection xmlParts = workbook.CustomXmlParts;
            for (int i = 0; i < xmlParts.Count; i++)
            {
                CustomXmlPart part = xmlParts[i];
                if (part != null && part.ID == RequiredCustomXmlPartId)
                {
                    containsRequiredPart = true;
                    break;
                }
            }

            // Process only if the required custom XML part is present
            if (containsRequiredPart)
            {
                // Example modification: add a new worksheet and write a note
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet newSheet = workbook.Worksheets[newSheetIndex];
                newSheet.Name = "Processed";
                newSheet.Cells["A1"].PutValue("Workbook contained the required custom XML part and was processed.");

                // Save the modified workbook (using the standard Save method as per lifecycle rules)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain the required custom XML part. No changes were made.");
            }

            // Clean up
            workbook.Dispose();
        }
    }
}
