using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyRangeWithUpdatedExternalRefs
{
    class Program
    {
        static void Main()
        {
            // Paths for source and destination workbooks
            string sourcePath = "Source.xlsx";
            string destinationPath = "Destination.xlsx";

            try
            {
                // Ensure source file exists; if not, create a simple workbook for demo purposes
                if (!File.Exists(sourcePath))
                {
                    var tempWb = new Workbook();
                    var ws = tempWb.Worksheets[0];
                    ws.Name = "Sheet1";
                    ws.Cells["A1"].PutValue("Demo");
                    ws.Cells["B2"].Formula = $"='[{Path.GetFileName(sourcePath)}]Sheet1'!A1";
                    tempWb.Save(sourcePath);
                }

                // Load the source workbook (create rule)
                Workbook sourceWb = new Workbook(sourcePath);

                // Create a new (empty) destination workbook (create rule)
                Workbook destWb = new Workbook();
                // Remove the default sheet to start clean
                destWb.Worksheets.Clear();

                // Add a worksheet to the destination workbook
                Worksheet destSheet = destWb.Worksheets.Add("Sheet1");

                // Define the range to copy from the source workbook
                // Example: copy cells A1:C5 from the first worksheet
                Aspose.Cells.Range srcRange = sourceWb.Worksheets[0].Cells.CreateRange("A1:C5");

                // Define the target range in the destination worksheet
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:C5");

                // Set paste options – copy everything (values, formulas, formats, etc.)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                    // IgnoreLinksToOriginalFile remains false to retain external links
                };

                // Copy the range using the provided Copy method with PasteOptions
                destRange.Copy(srcRange, pasteOptions);

                // Update any external references in formulas to point to the new file
                string srcFileName = Path.GetFileName(sourcePath);
                string destFileName = Path.GetFileName(destinationPath);

                for (int row = 0; row < destRange.RowCount; row++)
                {
                    for (int col = 0; col < destRange.ColumnCount; col++)
                    {
                        Cell cell = destRange[row, col];
                        if (cell.IsFormula)
                        {
                            string updatedFormula = cell.Formula.Replace(srcFileName, destFileName, StringComparison.OrdinalIgnoreCase);
                            cell.Formula = updatedFormula;
                        }
                    }
                }

                // Save the destination workbook (save rule)
                destWb.Save(destinationPath);

                Console.WriteLine("Range copied and external references updated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}