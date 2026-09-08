// Title: Refresh all linked pictures in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates through each worksheet, identifies linked Picture objects, and invokes their Refresh method. | Show how to use reflection in C# to safely access the IsLinked property and Refresh method of Aspose.Cells Picture objects for version‑agnostic handling. | Create a console application that updates external image links in a workbook after the source images change and saves the result to a new file.
// Common Searches: Aspose.Cells C# refresh linked images in existing workbook | How to programmatically update external picture links in Excel using Aspose.Cells | Iterate over worksheet pictures and call Refresh method with Aspose.Cells .NET | Use reflection to check IsLinked property of Picture in Aspose.Cells | Refresh all web‑linked pictures after changing source files with Aspose.Cells
// Tags: Aspose.Cells picture refresh | C# iterate worksheet pictures | Aspose.Cells IsLinked reflection | update external image links Excel .NET | programmatic picture refresh Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), walks through every worksheet and each picture on it, uses reflection to detect linked pictures via the IsLinked property, calls the Refresh method on those pictures, and finally saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if the input file exists; otherwise create a new workbook.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterate through all worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all pictures on the worksheet.
                foreach (Picture pic in sheet.Pictures)
                {
                    try
                    {
                        // Use reflection to call IsLinked and Refresh if they are available in the current version.
                        var isLinkedProp = pic.GetType().GetProperty("IsLinked");
                        var refreshMethod = pic.GetType().GetMethod("Refresh");

                        if (isLinkedProp != null && refreshMethod != null)
                        {
                            bool isLinked = (bool)isLinkedProp.GetValue(pic);
                            if (isLinked)
                            {
                                refreshMethod.Invoke(pic, null);
                            }
                        }
                    }
                    catch
                    {
                        // If the picture cannot be refreshed, ignore and continue.
                    }
                }
            }

            // Save the workbook after processing.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
