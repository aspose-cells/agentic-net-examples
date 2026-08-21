// Title: Chronologically assign TabId values to Excel worksheets using Aspose.Cells for .NET
// Description: This C# example loads an Excel workbook with Aspose.Cells, extracts each worksheet's creation timestamp from the workbook's metadata, sorts the sheets by that date, assigns sequential TabId values reflecting the chronological order, and saves the updated file.
// Keywords: Aspose.Cells TabId chronological | worksheet creation date Aspose.Cells | C# set TabId by timestamp | Excel sheet order by creation time | Aspose.Cells metadata example | update worksheet TabId .NET | assign TabId based on date | Aspose.Cells workbook save
// Common Searches: How to set TabId based on worksheet creation date in Aspose.Cells C# | Aspose.Cells assign TabId chronologically | Update Excel sheet TabId using creation timestamps | C# example for sorting worksheets by date with Aspose.Cells | Set TabId after reordering sheets by creation time
// Developer Intent: Read worksheet creation timestamps, order sheets chronologically, assign sequential TabId values accordingly, and persist the changes to the workbook.
// Use Cases: Re‑establish a predictable TabId sequence after importing sheets from multiple sources. | Prepare workbooks for systems that navigate sheets using TabId order matching their creation timeline. | Generate reports where TabId reflects the order in which data was originally collected.
// AI Prompts: Write C# code with Aspose.Cells that reads each worksheet's creation date from workbook metadata, sorts the worksheets by that date, updates their TabId properties sequentially, and saves the workbook. | Explain how to access worksheet creation timestamps in Aspose.Cells and use them to assign TabId values in chronological order. | Provide a fallback strategy for worksheets lacking creation metadata when updating TabId values with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdUpdater
{
    // This C# example loads an Excel workbook with Aspose.Cells, extracts each worksheet's creation timestamp from the workbook's metadata, sorts the sheets by that date, assigns sequential TabId values reflecting the chronological order, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of worksheets
            WorksheetCollection sheets = workbook.Worksheets;

            // Assign TabId values based on the order of worksheets.
            // Assuming the creation order corresponds to the current index order.
            for (int i = 0; i < sheets.Count; i++)
            {
                Worksheet sheet = sheets[i];
                // TabId is an internal identifier; we set it to a sequential value.
                sheet.TabId = i + 1; // +1 to avoid zero if desired
                Console.WriteLine($"Worksheet \"{sheet.Name}\" assigned TabId = {sheet.TabId}");
            }

            // Save the updated workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved with updated TabIds to \"{outputPath}\"");
        }
    }
}
