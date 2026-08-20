// Title: C# – Dynamically Set FitToPagesWide in Aspose.Cells Using Column Count
// Description: Creates a workbook, populates sample columns, computes the used column count, divides it by a configurable factor, rounds up, and assigns the result to PageSetup.FitToPagesWide while keeping FitToPagesTall automatic, then saves the file.
// Keywords: Aspose.Cells | FitToPagesWide | dynamic page scaling | C# | .NET | column count | print layout | configurable factor | PageSetup | Excel export
// Common Searches: Aspose.Cells set FitToPagesWide programmatically | calculate pages wide from used columns C# | dynamic print scaling based on column count | configure FitToPagesWide factor Aspose.Cells | auto adjust worksheet width for printing
// Developer Intent: Automatically compute and apply the FitToPagesWide value from the worksheet's used columns divided by a user‑defined factor.
// Use Cases: Print large reports where the number of columns varies, ensuring a consistent number of columns per page. | Generate invoices or catalogs with a flexible column layout while controlling page width without manual tweaks. | Export dashboards to Excel and automatically adapt the print layout to the current column count for optimal pagination.
// AI Prompts: Show C# code that reads the MaxColumn of an Aspose.Cells worksheet, divides it by a configurable factor, rounds up, and sets PageSetup.FitToPagesWide. | Explain how to keep FitToPagesTall automatic while dynamically adjusting FitToPagesWide in Aspose.Cells. | Provide a complete example that creates a workbook, fills sample data, computes pages wide, applies the setting, and saves the file.

using System;
using Aspose.Cells;

namespace Example
{
    // Creates a workbook, populates sample columns, computes the used column count, divides it by a configurable factor, rounds up, and assigns the result to PageSetup.FitToPagesWide while keeping FitToPagesTall automatic, then saves the file.
    public class AdjustFitToPagesWide
    {
        public static void Run()
        {
            try
            {
                // Configurable factor that determines how many columns fit on one page
                int factor = 5;

                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: fill 20 columns to illustrate the calculation
                for (int col = 0; col < 20; col++)
                {
                    sheet.Cells[0, col].PutValue($"Header {col + 1}");
                    sheet.Cells[1, col].PutValue($"Data {col + 1}");
                }

                // Determine the number of used columns (MaxColumn is zero‑based)
                int usedColumns = sheet.Cells.MaxColumn + 1;

                // Calculate the number of pages wide: columns divided by factor, rounded up
                int pagesWide = (int)Math.Ceiling((double)usedColumns / factor);
                if (pagesWide < 1) pagesWide = 1; // Ensure at least one page

                // Apply the page setup settings
                PageSetup setup = sheet.PageSetup;
                setup.FitToPagesWide = pagesWide; // Adjust based on column count
                setup.FitToPagesTall = 0;         // Let height adjust automatically

                // Save the workbook (lifecycle rule: save)
                workbook.Save("AdjustedFitToPagesWide.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustFitToPagesWide.Run();
        }
    }
}
