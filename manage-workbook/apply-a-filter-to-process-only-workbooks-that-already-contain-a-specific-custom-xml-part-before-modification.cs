using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlFilterDemo
{
    class Program
    {
        // The GUID of the custom XML part we are interested in.
        private const string TargetCustomXmlPartId = "2F087CB2-7CA8-43DA-B048-2E2F61F4936F";

        static void Main()
        {
            // Path to the source workbook.
            string sourcePath = "source.xlsx";
            // Path where the (potentially) modified workbook will be saved.
            string outputPath = "output.xlsx";

            // Load the workbook normally.
            Workbook wb = new Workbook(sourcePath);

            // Try to retrieve the custom XML part by its ID.
            CustomXmlPart part = wb.CustomXmlParts.SelectByID(TargetCustomXmlPartId);

            if (part != null)
            {
                // The workbook contains the required custom XML part.
                // Perform the desired modification – for example, add a new worksheet.
                int newSheetIndex = wb.Worksheets.Add();
                Worksheet newSheet = wb.Worksheets[newSheetIndex];
                newSheet.Name = "ProcessedSheet";

                // Write a note indicating that the workbook was processed.
                newSheet.Cells["A1"].PutValue("Workbook contained the required custom XML part and was processed.");

                Console.WriteLine("Custom XML part found. Workbook modified.");
            }
            else
            {
                // The required custom XML part is missing; skip modification.
                Console.WriteLine("Custom XML part not found. Workbook left unchanged.");
            }

            // Save the workbook (modified or unchanged) to the output path.
            wb.Save(outputPath);
        }
    }
}