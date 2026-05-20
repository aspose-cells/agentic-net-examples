using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class ProcessWorkbookWithCustomXmlPart
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // The ID of the custom XML part that must exist for processing
        string requiredPartId = "2F087CB2-7CA8-43DA-B048-2E2F61F4936F";

        // Load the workbook (using default LoadOptions)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Check if the workbook contains the required custom XML part
        CustomXmlPart requiredPart = workbook.CustomXmlParts.SelectByID(requiredPartId);

        if (requiredPart != null)
        {
            // The required custom XML part exists – proceed with modifications

            // Example modification: add a new worksheet and write a message
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "Processed";
            newSheet.Cells["A1"].PutValue("Workbook processed because required XML part was found.");

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        else
        {
            // Required custom XML part not found – skip processing
            Console.WriteLine("The workbook does not contain the required custom XML part. No changes were made.");
        }

        // Clean up
        workbook.Dispose();
    }
}