using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsBackgroundPdfExport
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths to the source Excel file and background image
                string sourceExcelPath = "input.xlsx";
                string backgroundImagePath = "background.jpg";

                // Verify that the required files exist
                if (!File.Exists(sourceExcelPath))
                {
                    Console.WriteLine($"Source Excel file not found: {sourceExcelPath}");
                    return;
                }

                if (!File.Exists(backgroundImagePath))
                {
                    Console.WriteLine($"Background image file not found: {backgroundImagePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourceExcelPath);

                // Read the background image into a byte array
                byte[] backgroundImageData = File.ReadAllBytes(backgroundImagePath);

                // Process each worksheet
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    try
                    {
                        Worksheet sheet = sourceWorkbook.Worksheets[i];

                        // Apply the background image to the worksheet
                        sheet.BackgroundImage = backgroundImageData;

                        // Create a temporary workbook containing only this worksheet
                        Workbook tempWorkbook = new Workbook();
                        tempWorkbook.Worksheets.Clear(); // Remove the default sheet

                        // Copy the current worksheet into the temporary workbook
                        // AddCopy expects the source sheet name, not the Worksheet object
                        tempWorkbook.Worksheets.AddCopy(sheet.Name);

                        // Define temporary file paths
                        string tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_sheet_{i}.xlsx");
                        string outputPdfPath = $"Sheet_{i + 1}.pdf";

                        // Save the temporary workbook (required for ConversionUtility)
                        tempWorkbook.Save(tempExcelPath);

                        // Convert the temporary Excel file to PDF
                        ConversionUtility.Convert(tempExcelPath, outputPdfPath);

                        Console.WriteLine($"Worksheet '{sheet.Name}' exported to PDF: {outputPdfPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing worksheet index {i}: {ex.Message}");
                    }
                    finally
                    {
                        // Clean up the temporary Excel file if it exists
                        string tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_sheet_{i}.xlsx");
                        if (File.Exists(tempExcelPath))
                        {
                            try
                            {
                                File.Delete(tempExcelPath);
                            }
                            catch (Exception delEx)
                            {
                                Console.WriteLine($"Failed to delete temporary file '{tempExcelPath}': {delEx.Message}");
                            }
                        }
                    }
                }

                Console.WriteLine("All worksheets have been processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}