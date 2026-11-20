// Title: Save an Excel workbook as Flat ODS (FODS) using Aspose.Cells OdsSaveOptions in C#
// Common Searches: C# code example for exporting Excel to Flat ODS with Aspose.Cells | How to use OdsSaveOptions to generate a .fods file in .NET | Aspose.Cells SaveFormat.FODS tutorial for C# developers | Convert workbook to flat ODS format using Aspose.Cells API | Generate Flat ODS file from worksheet data in C#
// Tags: Aspose.Cells OdsSaveOptions FODS export | C# flat ODS generation with Aspose | Save workbook as .fods using SaveFormat.FODS | Export Excel to Flat ODS via Aspose.Cells API | OdsSaveOptions configuration for Flat ODS

using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a new Workbook, adds sample data, configures OdsSaveOptions with SaveFormat.FODS, and saves the workbook as a Flat ODS (.fods) file using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Configure ODS save options to generate Flat ODS (FODS)
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.FODS);

        // Save the workbook as a FODS file
        workbook.Save("output.fods", saveOptions);
    }
}
