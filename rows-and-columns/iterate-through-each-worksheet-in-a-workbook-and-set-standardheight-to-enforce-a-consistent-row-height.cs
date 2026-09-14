// Title: Set a consistent StandardHeight for every worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets in a workbook and assigns a specific StandardHeight value to each sheet. | Show how to define a default row height, apply it to every worksheet, and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells C# set default row height for all worksheets in a workbook | How to apply the same StandardHeight to multiple sheets using Aspose.Cells | Iterate over workbook worksheets and change row height with Aspose.Cells C# | Saving an Excel file with uniform row height using Aspose.Cells | C# example for setting StandardHeight property on each worksheet
// Tags: Aspose.Cells set worksheet default row height | C# iterate workbook worksheets StandardHeight | apply uniform row height to multiple Excel sheets Aspose.Cells | save workbook with consistent row height Aspose.Cells | StandardHeight property usage Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds extra worksheets, sets the StandardHeight of every sheet to a defined value (e.g., 20 points), and saves the file as AllSheetsStandardHeight.xlsx.
    public class SetStandardRowHeightForAllSheets
    {
        public static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets to demonstrate iteration over multiple sheets
            workbook.Worksheets.Add(); // second worksheet
            workbook.Worksheets.Add(); // third worksheet

            // Define the consistent row height (in points) to apply to every worksheet
            double consistentHeight = 20.0;

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the default (standard) row height for the current worksheet
                sheet.Cells.StandardHeight = consistentHeight;
            }

            // Save the workbook to persist the changes
            workbook.Save("AllSheetsStandardHeight.xlsx", SaveFormat.Xlsx);
        }
    }
}
