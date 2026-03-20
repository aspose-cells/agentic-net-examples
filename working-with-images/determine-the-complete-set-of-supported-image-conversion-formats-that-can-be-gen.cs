using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "sample.xlsx";

        // If the sample file does not exist, create a simple workbook and save it (create‑save rule)
        if (!File.Exists(sourcePath))
        {
            Workbook wbCreate = new Workbook();                         // create workbook
            wbCreate.Worksheets[0].Cells["A1"].PutValue("Sample");     // add data
            wbCreate.Save(sourcePath, SaveFormat.Xlsx);                // save workbook as XLSX
        }

        // Load the workbook from the XLSX file (load rule)
        Workbook workbook = new Workbook(sourcePath);

        // ------------------------------------------------------------
        // 1. Image formats supported by the rendering engine (ImageType enum)
        // ------------------------------------------------------------
        Console.WriteLine("Supported ImageType formats for rendering:");
        foreach (ImageType imgType in Enum.GetValues(typeof(ImageType)))
        {
            if (imgType != ImageType.Unknown) // skip the placeholder value
            {
                Console.WriteLine("- " + imgType);
            }
        }

        // ------------------------------------------------------------
        // 2. Image formats supported by ImageSaveOptions (SaveFormat values)
        // ------------------------------------------------------------
        Console.WriteLine("\nSupported SaveFormat image formats (ImageSaveOptions):");
        SaveFormat[] imageSaveFormats = new SaveFormat[]
        {
            SaveFormat.Tiff,
            SaveFormat.Svg,
            SaveFormat.Bmp,
            SaveFormat.Png,
            SaveFormat.Jpg,
            SaveFormat.Emf,
            SaveFormat.Gif
        };
        foreach (SaveFormat fmt in imageSaveFormats)
        {
            Console.WriteLine("- " + fmt);
        }

        // ------------------------------------------------------------
        // 3. Demonstrate rendering the whole workbook to each image type
        //    using WorkbookRender and ImageOrPrintOptions (render‑to‑image rule)
        // ------------------------------------------------------------
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        foreach (ImageType imgType in Enum.GetValues(typeof(ImageType)))
        {
            if (imgType == ImageType.Unknown) continue; // skip unknown

            options.ImageType = imgType;                     // set desired image type
            WorkbookRender renderer = new WorkbookRender(workbook, options); // create renderer

            // Build output file name based on the image type
            string outFile = $"output_{imgType}.{imgType.ToString().ToLower()}";

            // Render the entire workbook to an image file (ToImage(string) rule)
            renderer.ToImage(outFile);

            Console.WriteLine($"Rendered workbook to {outFile}");
        }
    }
}