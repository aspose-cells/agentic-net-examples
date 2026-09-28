// Title: Set a custom worksheet paper size in millimeters by converting to points with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that receives width and height in millimeters, converts them to points (using 72 pts per inch and 25.4 mm per inch), and applies a custom paper size to the first worksheet via Aspose.Cells PageSetup. | Show how to call the method to apply an A5 (148 mm × 210 mm) page size to a newly created workbook and save the file as XLSX. | Demonstrate loading an existing XLSX file, changing its page setup to a 100 mm × 150 mm custom size, and saving the modified workbook.
// Common Searches: Aspose.Cells C# set custom page size using millimeters | convert millimeters to points for worksheet page setup Aspose.Cells | C# example custom A5 paper size Aspose.Cells | how to use PageSetup.CustomPaperSize method in Aspose.Cells | set custom page dimensions for first worksheet Aspose.Cells .NET
// Tags: custom page dimensions Aspose.Cells | millimeter to point conversion C# | PageSetup.CustomPaperSize method | worksheet page size .NET | A5 paper size Aspose.Cells

using System;
using System.Drawing; // For SizeF
using System.IO;      // For File.Exists
using Aspose.Cells;   // Aspose.Cells namespace

namespace Example
{
    // Provides a C# helper that converts width and height from millimeters to points (72/25.4 factor) and assigns a custom page size to the first worksheet using PageSetup.PaperSize = PaperSizeType.Custom and PageSetup.CustomPaperSize(widthPoints, heightPoints). Includes examples for creating a new workbook with A5 size and modifying an existing file's page setup.
    public class PaperSizeHelper
    {
        /// <param name="workbook">The Aspose.Cells workbook to modify.</param>
        /// <param name="widthMm">Paper width in millimeters.</param>
        /// <param name="heightMm">Paper height in millimeters.</param>
        public static void SetCustomPaperSize(Workbook workbook, double widthMm, double heightMm)
        {
            // 1 inch = 25.4 mm, 1 point = 1/72 inch
            const double MmToPoints = 72.0 / 25.4;

            // Convert millimeters to points
            float widthPoints = (float)(widthMm * MmToPoints);
            float heightPoints = (float)(heightMm * MmToPoints);

            // Access the PageSetup of the first worksheet
            PageSetup pageSetup = workbook.Worksheets[0].PageSetup;

            // Use a custom paper size
            pageSetup.PaperSize = PaperSizeType.Custom;

            // Assign the custom size (width, height) in points.
            // In some Aspose.Cells versions CustomPaperSize is a method; use it accordingly.
            pageSetup.CustomPaperSize(widthPoints, heightPoints);
        }
    }

    class Program
    {
        static void Main()
        {
            // Example: create a new workbook and set custom paper size
            try
            {
                Workbook wbCreate = new Workbook();
                wbCreate.Worksheets[0].Cells["A1"].PutValue("Sample");
                PaperSizeHelper.SetCustomPaperSize(wbCreate, 148, 210); // A5 size
                wbCreate.Save("CustomPaper_Create.xlsx");
                Console.WriteLine("Workbook created and saved as 'CustomPaper_Create.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during workbook creation: {ex.Message}");
            }

            // Example: load an existing workbook, modify paper size, and save
            try
            {
                const string inputPath = "ExistingFile.xlsx";

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Skipping load example.");
                }
                else
                {
                    Workbook wbLoad = new Workbook(inputPath);
                    PaperSizeHelper.SetCustomPaperSize(wbLoad, 100, 150);
                    wbLoad.Save("CustomPaper_Loaded.xlsx");
                    Console.WriteLine("Modified workbook saved as 'CustomPaper_Loaded.xlsx'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during workbook load/save: {ex.Message}");
            }
        }
    }
}
