// Title: Detect localization differences between English and Chinese Excel chart images using Aspose.Cells in C#
// Common Searches: how to export an Excel chart to PNG with Aspose.Cells C# | compare two chart images byte by byte in C# | verify language‑specific chart rendering differences in Excel workbooks | programmatic localization testing of Excel charts using Aspose.Cells | Aspose.Cells chart ToImage example for detecting image changes
// Tags: Aspose.Cells chart PNG export | byte array image diff detection C# | Excel chart localization validation Aspose.Cells | chart rendering consistency across languages | programmatic chart image verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType enum

// The sample loads EnglishChart.xlsx and ChineseChart.xlsx, extracts the first chart from each workbook, renders each chart to a PNG byte array using Aspose.Cells ImageOrPrintOptions, compares the byte arrays to identify any differences, reports whether the images are identical, and saves the PNG files for visual inspection when discrepancies are found.
class Program
{
    static void Main(string[] args)
    {
        const string englishWorkbookPath = "EnglishChart.xlsx";
        const string chineseWorkbookPath = "ChineseChart.xlsx";

        // Validate input files
        if (!File.Exists(englishWorkbookPath))
        {
            Console.WriteLine($"Error: File not found -> {englishWorkbookPath}");
            return;
        }
        if (!File.Exists(chineseWorkbookPath))
        {
            Console.WriteLine($"Error: File not found -> {chineseWorkbookPath}");
            return;
        }

        try
        {
            // Load both workbooks
            Workbook englishWb = new Workbook(englishWorkbookPath);
            Workbook chineseWb = new Workbook(chineseWorkbookPath);

            // Assume the chart we want to compare is the first chart on the first worksheet of each workbook
            Chart englishChart = englishWb.Worksheets[0].Charts[0];
            Chart chineseChart = chineseWb.Worksheets[0].Charts[0];

            // Export charts to PNG byte arrays (in-memory)
            byte[] englishImg = ChartToPngBytes(englishChart);
            byte[] chineseImg = ChartToPngBytes(chineseChart);

            // Compare the two images byte‑by‑byte
            bool areIdentical = CompareByteArrays(englishImg, chineseImg, out int diffByteCount);

            // Output result
            if (areIdentical)
            {
                Console.WriteLine("The English and Chinese chart images are identical. No localization differences detected.");
            }
            else
            {
                Console.WriteLine($"The chart images differ. Different byte count: {diffByteCount}");

                // Save the two images for visual inspection
                const string englishImgPath = "EnglishChart.png";
                const string chineseImgPath = "ChineseChart.png";

                File.WriteAllBytes(englishImgPath, englishImg);
                File.WriteAllBytes(chineseImgPath, chineseImg);

                Console.WriteLine($"English chart image saved to: {englishImgPath}");
                Console.WriteLine($"Chinese chart image saved to: {chineseImgPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An exception occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    private static byte[] ChartToPngBytes(Chart chart)
    {
        // Configure image options for PNG output
        ImageOrPrintOptions opts = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            OnePagePerSheet = true
        };

        try
        {
            using (MemoryStream ms = new MemoryStream())
            {
                // Render chart directly to the stream
                chart.ToImage(ms, opts);
                return ms.ToArray();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to convert chart to PNG bytes: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    private static bool CompareByteArrays(byte[] arr1, byte[] arr2, out int diffByteCount)
    {
        diffByteCount = 0;

        if (arr1 == null || arr2 == null)
        {
            diffByteCount = -1; // sentinel for null data
            return false;
        }

        if (arr1.Length != arr2.Length)
        {
            diffByteCount = -1; // sentinel for size mismatch
            return false;
        }

        for (int i = 0; i < arr1.Length; i++)
        {
            if (arr1[i] != arr2[i])
                diffByteCount++;
        }

        return diffByteCount == 0;
    }
}
