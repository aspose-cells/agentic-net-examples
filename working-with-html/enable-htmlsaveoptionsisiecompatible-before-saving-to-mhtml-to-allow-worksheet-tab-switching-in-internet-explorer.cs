using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMhtmlExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Ensure the first worksheet exists (index 0 is always present in a new workbook)
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("First sheet");

                // Add a second worksheet and obtain its reference
                int sheet2Index = workbook.Worksheets.Add();          // Returns the index of the new sheet
                Worksheet sheet2 = workbook.Worksheets[sheet2Index];
                sheet2.Name = "Sheet2";
                sheet2.Cells["A1"].PutValue("Second sheet");

                // Create HTML save options for MHTML format
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml)
                {
                    // Enable IE compatibility to allow worksheet tab switching in Internet Explorer
                    IsIECompatible = true
                };

                // Define output file path
                string outputPath = "WorkbookWithTabs.mht";

                // Save the workbook as MHTML using the save options
                workbook.Save(outputPath, saveOptions);
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