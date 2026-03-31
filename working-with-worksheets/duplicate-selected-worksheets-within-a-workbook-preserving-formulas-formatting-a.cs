using System;
using Aspose.Cells;

namespace WorksheetDuplicationDemo
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            string sourceFile = "input.xlsx";
            Workbook workbook = new Workbook(sourceFile);

            // List of worksheet names that need to be duplicated
            string[] sheetsToDuplicate = { "Sheet1", "Data", "Summary" };

            foreach (string sheetName in sheetsToDuplicate)
            {
                // Verify that the source worksheet exists
                if (workbook.Worksheets[sheetName] == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found. Skipping.");
                    continue;
                }

                // Add a copy of the worksheet to the collection.
                // AddCopy returns the index of the newly created worksheet.
                int copiedIndex = workbook.Worksheets.AddCopy(sheetName);

                // Rename the copied worksheet to indicate it is a duplicate.
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = $"{sheetName}_Copy";
            }

            // Save the workbook with duplicated worksheets.
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);

            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
    }
}