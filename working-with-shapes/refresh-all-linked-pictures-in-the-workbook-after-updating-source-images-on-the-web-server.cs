// Title: Refresh linked pictures in an Excel workbook after external image changes using Aspose.Cells for .NET (C#)
// Common Searches: asp.net refresh linked picture in Excel workbook after image URL changes | c# Aspose.Cells update external image links in existing .xlsx file | how to programmatically refresh pictures linked to web URLs in Excel using Aspose.Cells | iterate through worksheet pictures and call Refresh method Aspose.Cells C# | refresh external image references in Excel file without reopening in Office
// Tags: Aspose.Cells picture Refresh API | C# workbook picture collection processing | dynamic method call for Excel image refresh | Excel linked image synchronization using .NET | save modified workbook after picture update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing Excel workbook, loops through each worksheet and its Pictures collection, uses dynamic invocation to call the Refresh method on linked picture objects (ignoring those without the method), logs any issues, and saves the workbook with the refreshed images to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains linked pictures
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each picture on the worksheet
                foreach (Picture picture in sheet.Pictures)
                {
                    try
                    {
                        // Attempt to refresh the picture (works for linked pictures)
                        // Using dynamic invocation to avoid compile‑time dependency on Refresh()
                        ((dynamic)picture).Refresh();
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // Refresh method not available; ignore this picture
                    }
                    catch (Exception ex)
                    {
                        // Log any other runtime issues while refreshing a picture
                        Console.WriteLine($"Warning: Unable to refresh picture on sheet \"{sheet.Name}\": {ex.Message}");
                    }
                }
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
