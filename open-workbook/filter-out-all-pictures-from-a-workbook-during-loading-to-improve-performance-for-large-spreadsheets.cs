// Title: C# – Load an Excel workbook without pictures using Aspose.Cells for faster performance
// Description: Demonstrates how to create a LoadOptions object, apply a LoadFilter that excludes picture data (LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture), open a large workbook, verify that the Pictures collection is empty, and save the file without any embedded images. This reduces memory usage and speeds up processing of big spreadsheets.
// Keywords: Aspose.Cells load workbook without pictures | C# exclude images Excel load | LoadFilter picture flag | Improve Excel loading performance | LoadDataFilterOptions picture removal | memory efficient Excel processing
// Common Searches: how to skip pictures when loading Excel with Aspose.Cells | load large workbook without images C# | Aspose.Cells LoadFilter exclude pictures | optimize Excel file loading performance .NET | remove picture objects during workbook load
// Developer Intent: Open an Excel file while omitting all picture objects to lower memory consumption and accelerate processing.
// Use Cases: Run calculations on massive financial reports where only cell data matters. | Generate data extracts on low‑memory servers without the overhead of embedded graphics. | Perform batch data transformations on large workbooks in cloud services where image data is unnecessary.
// AI Prompts: Show me C# code to load an Excel workbook with Aspose.Cells while ignoring all pictures and then save it. | Provide an example of using LoadDataFilterOptions to exclude pictures when opening a workbook in .NET. | Explain how to confirm that no pictures were loaded after applying a LoadFilter in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a LoadOptions object, apply a LoadFilter that excludes picture data (LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture), open a large workbook, verify that the Pictures collection is empty, and save the file without any embedded images. This reduces memory usage and speeds up processing of big spreadsheets.
class FilterPicturesDemo
{
    static void Main()
    {
        // Input and output file paths
        string inputFile = "largeWorkbook.xlsx";
        string outputFile = "largeWorkbook_NoPictures.xlsx";

        // Create LoadOptions instance
        LoadOptions loadOptions = new LoadOptions();

        // Configure LoadFilter to load everything except pictures
        // LoadDataFilterOptions.All includes all data; we remove the Picture flag using bitwise AND NOT
        LoadDataFilterOptions filter = LoadDataFilterOptions.All & ~LoadDataFilterOptions.Picture;
        loadOptions.LoadFilter = new LoadFilter(filter);

        // Load the workbook with the configured options
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Verify that pictures are not loaded (should be 0)
        Console.WriteLine("Pictures count after load: " + workbook.Worksheets[0].Pictures.Count);

        // Save the workbook; it will be saved without any picture objects
        workbook.Save(outputFile, SaveFormat.Xlsx);
    }
}
