using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace WorksheetIdValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Rename the original sheet and add some data
            Worksheet original = workbook.Worksheets[0];
            original.Name = "Original";
            original.Cells["A1"].PutValue("Data in original sheet");

            // Add a second worksheet
            Worksheet second = workbook.Worksheets.Add("Second");
            second.Cells["A1"].PutValue("Data in second sheet");

            // Copy the "Original" worksheet using AddCopy (by name)
            int copiedIndex = workbook.Worksheets.AddCopy("Original");
            Worksheet copied = workbook.Worksheets[copiedIndex];
            copied.Name = "Copied";

            // Validate that each worksheet has a unique TabId (internal sheet identifier)
            HashSet<int> tabIds = new HashSet<int>();
            bool allUnique = true;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int id = sheet.TabId; // internal identifier for the sheet
                if (!tabIds.Add(id))
                {
                    // Duplicate found
                    Console.WriteLine($"Duplicate TabId detected: {id} on sheet \"{sheet.Name}\"");
                    allUnique = false;
                }
            }

            if (allUnique)
            {
                Console.WriteLine("All worksheets have unique TabId values.");
            }

            // Save the workbook (output file)
            workbook.Save("WorksheetIdValidationResult.xlsx");
        }
    }
}