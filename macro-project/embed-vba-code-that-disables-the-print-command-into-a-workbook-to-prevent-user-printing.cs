using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaPrintDisableDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // The workbook already contains a class module named "ThisWorkbook"
                // Retrieve that existing module instead of adding a duplicate
                VbaModule thisWorkbookModule = workbook.VbaProject.Modules["ThisWorkbook"];

                // VBA code that cancels any print attempt
                // The Workbook_BeforePrint event is triggered before printing; setting Cancel = True stops the print
                string vbaCode =
                    "Private Sub Workbook_BeforePrint(Cancel As Boolean)\r\n" +
                    "    Cancel = True\r\n" +
                    "    MsgBox \"Printing is disabled for this workbook.\"\r\n" +
                    "End Sub";

                // Assign the VBA code to the module
                thisWorkbookModule.Codes = vbaCode;

                // (Optional) Protect the VBA project to prevent users from viewing/modifying the code
                // Here we protect without locking for viewing, using a password
                workbook.VbaProject.Protect(false, "vbaPassword123");

                // Save the workbook as a macro‑enabled file (.xlsm) so the VBA code is retained
                string outputPath = "WorkbookWithPrintDisabled.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved to '{outputPath}'. Printing is now disabled via VBA.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}