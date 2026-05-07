using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class WorkbookProcessing
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Input workbook files to be merged
        string[] sourceFiles = { "book1.xlsx", "book2.xlsx", "book3.xlsx" };

        // Output file names
        string mergedFile = "merged.xlsx";
        string pdfFile = "merged.pdf";

        // Load options for XLSX files
        LoadOptions loadOpts = new LoadOptions(LoadFormat.Xlsx);
        loadOpts.IgnoreNotPrinted = false;

        // Load the first workbook – this will become the master workbook
        Workbook master = new Workbook(sourceFiles[0], loadOpts);

        // Merge the remaining workbooks into the master workbook
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Workbook wb = new Workbook(sourceFiles[i], loadOpts);
            master.Combine(wb);
        }

        // Protect the workbook structure with a password
        master.Protect(ProtectionType.Structure, "structPwd");

        // Save PDF version of the merged workbook
        master.Save(pdfFile, SaveFormat.Pdf);

        // Save the merged workbook (protected structure)
        master.Save(mergedFile, SaveFormat.Xlsx);
    }
}