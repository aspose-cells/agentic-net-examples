// Title: Save a modified workbook to a new XLSX file while preserving all original formatting, styles, and textures with Aspose.Cells (C#)
// AI Prompts: Create C# code that loads an existing .xlsx, applies optional changes, and writes it to a different file using Aspose.Cells, ensuring that cell styles, themes, and embedded images remain unchanged. | Write a function that copies an Excel workbook to a new location with Aspose.Cells, preserving formatting, conditional formatting, and texture fills.
// Common Searches: asp.net c# aspose.cells save workbook to new file keep original cell styles | preserve background textures when saving modified Excel with Aspose.Cells | copy existing .xlsx and retain conditional formatting using Aspose.Cells C# | how to save workbook without losing embedded pictures aspose.cells | Aspose.Cells SaveFormat.Xlsx keep all formatting and images
// Tags: save workbook preserving formatting Aspose.Cells | copy Excel file retain cell textures C# | Aspose.Cells SaveFormat.Xlsx keep styles | preserve embedded images when saving workbook | modify and save workbook without losing conditional formatting

using Aspose.Cells;

// Loads an existing .xlsx with Aspose.Cells, optionally updates cells, and saves it as a new .xlsx file while preserving all formatting, styles, textures, and embedded images.
class Program
{
    static void Main()
    {
        // Load the original workbook (all formatting and textures are retained on load)
        Workbook workbook = new Workbook("original.xlsx");

        // ---- Place any modifications here ----
        // Example: workbook.Worksheets[0].Cells["A1"].PutValue("Updated value");
        // -------------------------------------

        // Save the workbook to a new file, preserving the original formatting and textures
        workbook.Save("modified.xlsx", SaveFormat.Xlsx);
    }
}
