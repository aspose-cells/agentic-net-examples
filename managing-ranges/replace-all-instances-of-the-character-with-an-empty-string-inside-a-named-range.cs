using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplaceInNamedRange
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data containing the '*' character
                sheet.Cells["A1"].PutValue("Hello*World");
                sheet.Cells["A2"].PutValue("Sample*Text");
                sheet.Cells["B1"].PutValue("NoAsterisk");
                sheet.Cells["B2"].PutValue("*LeadingAsterisk");
                sheet.Cells["C1"].PutValue("OutsideRange*");

                // Define a named range that covers A1:B2
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$2";

                // Retrieve the address part of the named range (e.g., "$A$1:$B$2")
                string refersTo = workbook.Worksheets.Names["MyRange"].RefersTo;
                int exclPos = refersTo.IndexOf('!');
                string address = refersTo.Substring(exclPos + 1);

                // Create the Aspose.Cells.Range object based on the address
                AsposeRange targetRange = sheet.Cells.CreateRange(address);

                // Iterate through each cell in the range and remove '*' characters
                foreach (Cell cell in targetRange)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        if (original.Contains("*"))
                        {
                            string updated = original.Replace("*", string.Empty);
                            cell.PutValue(updated);
                        }
                    }
                }

                // Define output file path
                string outputPath = "NamedRangeAsteriskRemoved.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}