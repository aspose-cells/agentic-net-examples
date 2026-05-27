using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeMapping
{
    class Program
    {
        static void Main()
        {
            // Configuration: map worksheet names to desired paper sizes
            var paperSizeMapping = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", PaperSizeType.PaperA4 },
                { "Sheet2", PaperSizeType.PaperLetter },
                { "Report", PaperSizeType.PaperLegal }
                // Add more mappings as needed
            };

            // Create LoadOptions (default settings)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the LoadOptions
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Apply the configured paper size to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMapping.TryGetValue(sheet.Name, out PaperSizeType size))
                {
                    // Set the paper size for the current worksheet
                    sheet.PageSetup.PaperSize = size;
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}