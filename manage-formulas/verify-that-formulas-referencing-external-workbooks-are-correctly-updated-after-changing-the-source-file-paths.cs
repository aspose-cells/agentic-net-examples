using System;
using Aspose.Cells;

namespace ExternalLinkUpdateDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // Step 1: Create the original external workbook (ExternalOld.xlsx)
            // -----------------------------------------------------------------
            Workbook externalOld = new Workbook();
            Worksheet extOldSheet = externalOld.Worksheets[0];
            extOldSheet.Name = "Sheet1";
            extOldSheet.Cells["A1"].PutValue("Old Value");
            externalOld.Save("ExternalOld.xlsx");

            // -----------------------------------------------------------------
            // Step 2: Create the main workbook that references the external file
            // -----------------------------------------------------------------
            Workbook mainWorkbook = new Workbook();
            Worksheet mainSheet = mainWorkbook.Worksheets[0];
            // Formula points to the old external workbook
            mainSheet.Cells["A1"].Formula = "=[ExternalOld.xlsx]Sheet1!A1";

            // Ensure the external link is registered (Aspose does this automatically)
            Console.WriteLine("Initial external link count: " + mainWorkbook.Worksheets.ExternalLinks.Count);

            // -----------------------------------------------------------------
            // Step 3: Change the external link path to a new file (ExternalNew.xlsx)
            // -----------------------------------------------------------------
            // Create the new external workbook with updated data
            Workbook externalNew = new Workbook();
            Worksheet extNewSheet = externalNew.Worksheets[0];
            extNewSheet.Name = "Sheet1";
            extNewSheet.Cells["A1"].PutValue("New Value");
            externalNew.Save("ExternalNew.xlsx");

            // Update the DataSource of the existing external link to point to the new file
            if (mainWorkbook.Worksheets.ExternalLinks.Count > 0)
            {
                ExternalLink link = mainWorkbook.Worksheets.ExternalLinks[0];
                // Optionally also update OriginalDataSource for completeness
                link.OriginalDataSource = link.OriginalDataSource?.Replace("ExternalOld.xlsx", "ExternalNew.xlsx");
                link.DataSource = "ExternalNew.xlsx";
                Console.WriteLine("External link updated to: " + link.DataSource);
            }

            // -----------------------------------------------------------------
            // Step 4: Refresh the linked data source and recalculate formulas
            // -----------------------------------------------------------------
            // Provide the new external workbook to UpdateLinkedDataSource
            mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalNew });

            // Recalculate formulas so that A1 reflects the new external value
            mainWorkbook.CalculateFormula();

            // -----------------------------------------------------------------
            // Step 5: Verify the updated value
            // -----------------------------------------------------------------
            string updatedValue = mainSheet.Cells["A1"].StringValue;
            Console.WriteLine("Updated value in main workbook cell A1: " + updatedValue);

            // -----------------------------------------------------------------
            // Step 6: Save the main workbook (optional)
            // -----------------------------------------------------------------
            mainWorkbook.Save("MainWorkbook_Updated.xlsx");
        }
    }
}