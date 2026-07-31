// Title: Aspose.Cells for .NET – Set 120% Zoom on All Worksheets and Export Each to a Separate XLS File
// Description: A C# example that creates a workbook, iterates through every worksheet, sets the Zoom property to 120 %, copies each sheet into its own workbook, removes the default sheet, and saves the result as an Excel 97‑2003 (.xls) file named after the original worksheet.
// Keywords: Aspose.Cells | C# worksheet zoom | set worksheet zoom .NET | export each sheet to separate xls | split workbook Aspose.Cells | Excel 97-2003 format | Zoom property Aspose | AddCopy worksheet | Workbook.Save XLS | Aspose.Cells example
// Common Searches: Aspose.Cells set worksheet zoom | C# save each worksheet as separate XLS | apply 120% zoom to all sheets using Aspose | split workbook into individual files Aspose.Cells | export worksheets to Excel 97-2003 format .NET | copy worksheet to new workbook Aspose.Cells
// Developer Intent: Apply a 120 % zoom to every worksheet and save each sheet as an individual XLS file.
// Use Cases: Create legacy Excel 97‑2003 files where every sheet opens at a uniform 120 % zoom for consistent on‑screen viewing. | Break a multi‑sheet report into separate workbooks, each pre‑configured with the required zoom level before distribution. | Automate generation of per‑sheet files for downstream systems that only accept single‑sheet XLS documents.
// AI Prompts: Generate C# code with Aspose.Cells that sets a 120% zoom on all worksheets and saves each as a separate .xls file. | Refactor the example to use MemoryStream for in‑memory saving of each workbook instead of writing to disk. | Explain best practices for handling exceptions when copying worksheets and assigning the Zoom property in Aspose.Cells. | Show how to customize the output file name pattern (e.g., include timestamps) while exporting each worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsZoomExample
{
    // A C# example that creates a workbook, iterates through every worksheet, sets the Zoom property to 120 %, copies each sheet into its own workbook, removes the default sheet, and saves the result as an Excel 97‑2003 (.xls) file named after the original worksheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a sample workbook with multiple worksheets
                Workbook sourceWorkbook = new Workbook();
                sourceWorkbook.Worksheets[0].Name = "Sheet1";
                sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");

                // Add additional sheets
                Worksheet sheet2 = sourceWorkbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].PutValue("Data in Sheet2");

                Worksheet sheet3 = sourceWorkbook.Worksheets.Add("Sheet3");
                sheet3.Cells["A1"].PutValue("Data in Sheet3");

                // Iterate through each worksheet in the source workbook
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    try
                    {
                        // Set the zoom level of the current worksheet to 120%
                        Worksheet currentSheet = sourceWorkbook.Worksheets[i];
                        currentSheet.Zoom = 120; // Zoom expects a percentage between 10 and 400

                        // Create a new workbook that will contain only this worksheet
                        Workbook singleSheetWorkbook = new Workbook();

                        // Remove the default sheet that comes with a new workbook
                        singleSheetWorkbook.Worksheets.Clear();

                        // Copy the current worksheet into the new workbook
                        // AddCopy(string sourceSheetName) creates a duplicate of the worksheet in the target collection
                        singleSheetWorkbook.Worksheets.AddCopy(currentSheet.Name);

                        // Prepare file name and ensure the directory exists
                        string fileName = $"Worksheet_{currentSheet.Name}.xls";
                        string directory = Path.GetDirectoryName(Path.GetFullPath(fileName));
                        if (!Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        // Save the new workbook as an XLS file (Excel 97-2003 format)
                        singleSheetWorkbook.Save(fileName, SaveFormat.Excel97To2003);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing sheet '{sourceWorkbook.Worksheets[i].Name}': {ex.Message}");
                    }
                }

                Console.WriteLine("All worksheets have been saved with a 120% zoom level.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
