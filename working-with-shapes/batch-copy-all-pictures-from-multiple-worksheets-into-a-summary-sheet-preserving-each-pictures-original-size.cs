using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class BatchCopyPicturesToSummary
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_with_summary.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Add or get summary worksheet
            Worksheet summarySheet = workbook.Worksheets["Summary"];
            if (summarySheet == null)
                summarySheet = workbook.Worksheets.Add("Summary");

            // Iterate through worksheets except summary
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws == summarySheet) continue;

                // Get all pictures from worksheet
                Picture[] pictures = ws.GetAllPictures();

                foreach (Picture srcPic in pictures)
                {
                    // Get image data; skip if none (e.g., linked picture)
                    byte[] imgData = srcPic.Data;
                    if (imgData == null || imgData.Length == 0) continue;

                    // Determine picture position
                    int topRow = srcPic.UpperLeftRow;
                    int leftColumn = srcPic.UpperLeftColumn;
                    int bottomRow = srcPic.LowerRightRow;
                    int rightColumn = srcPic.LowerRightColumn;

                    // Add picture to summary sheet
                    Picture targetPic;
                    using (MemoryStream ms = new MemoryStream(imgData))
                    {
                        int picIndex = summarySheet.Pictures.Add(topRow, leftColumn, bottomRow, rightColumn, ms);
                        targetPic = summarySheet.Pictures[picIndex];
                    }

                    // Copy properties from source picture
                    CopyOptions copyOptions = new CopyOptions();
                    targetPic.Copy(srcPic, copyOptions);
                }
            }

            // Save workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}