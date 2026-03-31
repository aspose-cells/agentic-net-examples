using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class IncludePdfLibraryInVba
{
    static void Main()
    {
        // Step 1: Create a new workbook (initially without VBA project)
        Workbook workbook = new Workbook();

        // Step 2: Save as macro-enabled workbook to create a VBA project
        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        workbook.Save(tempPath, SaveFormat.Xlsm);

        // Step 3: Reload the workbook so that the VBA project is available
        workbook = new Workbook(tempPath);

        // Clean up the temporary file
        File.Delete(tempPath);

        // Step 4: Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Step 5: Add a reference to the PDF library.
        // Replace the libid strings with the actual identifiers of the PDF COM library you need.
        // Example placeholders are used here.
        string referenceName = "PdfLibrary";
        string absoluteLibId = "*\\G{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}#1.0#0#C:\\Program Files\\PdfLibrary\\PdfLib.tlb#Pdf Library";
        string relativeLibId = "PdfLib.tlb#Pdf Library";

        // Use AddProjectRefrernce method as defined in the API
        vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibId, relativeLibId);

        // Step 6: Save the workbook with the new VBA reference
        string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WorkbookWithPdfReference.xlsm");
        workbook.Save(outputPath, SaveFormat.Xlsm);

        Console.WriteLine($"Workbook saved with PDF library reference at: {outputPath}");
    }
}