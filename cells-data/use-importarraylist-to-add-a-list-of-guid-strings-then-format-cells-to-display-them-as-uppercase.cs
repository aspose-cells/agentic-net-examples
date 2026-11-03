// Title: Import an ArrayList of GUID strings into an Excel column and format them as uppercase with Aspose.Cells for .NET
// Common Searches: how to import a C# ArrayList of GUIDs into an Excel worksheet using Aspose.Cells | Aspose.Cells C# convert imported string cells to uppercase | populate Excel column from ArrayList and apply uppercase formatting with Aspose.Cells | using ImportArrayList to add GUID values to Excel and change case in .NET | save workbook with uppercase GUIDs using Aspose.Cells C#
// Tags: ImportArrayList GUID Aspose.Cells | convert cell string to upper case Aspose.Cells | populate column from ArrayList Aspose.Cells | uppercase GUID display Aspose.Cells | save workbook GuidList.xlsx Aspose.Cells

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, generates an ArrayList of lowercase GUID strings, imports the list into the first column of the first worksheet with ImportArrayList, iterates over the imported cells to replace each string with its uppercase version, and saves the workbook as GuidList.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare an ArrayList of GUID strings (lowercase for demonstration)
            ArrayList guidList = new ArrayList();
            for (int i = 0; i < 5; i++)
            {
                guidList.Add(Guid.NewGuid().ToString().ToLower());
            }

            // Import the ArrayList into the worksheet starting at cell A1
            // Parameters: list, startRow, startColumn, isVertical (false = single column)
            sheet.Cells.ImportArrayList(guidList, 0, 0, false);

            // Convert each imported GUID to uppercase for display
            int rowCount = guidList.Count;
            for (int row = 0; row < rowCount; row++)
            {
                Cell cell = sheet.Cells[row, 0];
                if (cell.Type == CellValueType.IsString)
                {
                    cell.PutValue(cell.StringValue.ToUpper());
                }
            }

            // Define output file path
            string outputPath = "GuidList.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
