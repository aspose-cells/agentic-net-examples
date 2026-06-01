using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsZoomExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a sample workbook with multiple worksheets
                Workbook sourceWorkbook = new Workbook();
                sourceWorkbook.Worksheets[0].Name = "FirstSheet";
                sourceWorkbook.Worksheets.Add("SecondSheet");
                sourceWorkbook.Worksheets.Add("ThirdSheet");

                // Apply 120% zoom to every worksheet in the source workbook
                foreach (Worksheet ws in sourceWorkbook.Worksheets)
                {
                    ws.Zoom = 120; // Zoom is a percentage (10‑400)
                }

                // Save each worksheet as a separate XLS file
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    // Create a new workbook that will contain only the current worksheet
                    Workbook singleSheetWorkbook = new Workbook();

                    // Remove the default sheet created by the constructor
                    singleSheetWorkbook.Worksheets.Clear();

                    // Copy the current worksheet into the new workbook
                    // AddCopy expects the source sheet name, not the Worksheet object
                    singleSheetWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[i].Name);

                    // Ensure the zoom level is set (copy retains it, but set explicitly for safety)
                    singleSheetWorkbook.Worksheets[0].Zoom = 120;

                    // Build a file name for the sheet (e.g., FirstSheet.xls)
                    string fileName = $"{sourceWorkbook.Worksheets[i].Name}.xls";

                    // Delete existing file to avoid issues on overwrite
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }

                    // Save the workbook as an Excel 97‑2003 file
                    singleSheetWorkbook.Save(fileName, SaveFormat.Excel97To2003);
                }

                Console.WriteLine("All worksheets have been saved with 120% zoom.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}