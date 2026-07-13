using System;
using Aspose.Cells;

namespace AsposeCellsColumnFitValidation
{
    class Program
    {
        // Author: Aspose.Cells .NET example – validates column fitting per worksheet
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Flag indicating whether we want all columns to fit on one page
            bool FitAllColumnsOnOnePage = true;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (FitAllColumnsOnOnePage)
                {
                    PageSetup pageSetup = sheet.PageSetup;

                    // Validation: FitToPagesTall must be zero to allow all columns on a single page
                    if (pageSetup.FitToPagesTall != 0)
                    {
                        // If not set correctly, adjust it (or you could throw an exception)
                        pageSetup.FitToPagesTall = 0;
                    }

                    // Optionally, also ensure FitToPagesWide is zero to avoid width constraints
                    // (not required by the documentation but often used together)
                    if (pageSetup.FitToPagesWide != 0)
                    {
                        pageSetup.FitToPagesWide = 0;
                    }
                }
            }

            // Save the workbook after validation (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}