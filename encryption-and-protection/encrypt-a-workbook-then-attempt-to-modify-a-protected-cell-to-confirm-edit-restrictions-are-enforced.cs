using System;
using Aspose.Cells;

class WorkbookEncryptionAndProtectionDemo
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Original");

        // ---------- Encrypt the workbook with a password ----------
        wb.Settings.Password = "filePwd";
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // ---------- Define an editable range for cell A1 ----------
        // Add a protected range covering A1
        int rangeIndex = ws.AllowEditRanges.Add("EditableA1", 0, 0, 0, 0);
        ProtectedRange range = ws.AllowEditRanges[rangeIndex];
        // Set a password for this range
        range.Password = "rangePwd";

        // ---------- Protect the worksheet (all protection types) ----------
        ws.Protect(ProtectionType.All);

        // ---------- Save the encrypted and protected workbook ----------
        string filePath = "EncryptedProtectedWorkbook.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);

        // ---------- Load the workbook using the file password ----------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "filePwd";
        Workbook loadedWb = new Workbook(filePath, loadOptions);
        Worksheet loadedWs = loadedWb.Worksheets[0];

        // ---------- Attempt to modify the protected cell without providing the range password ----------
        try
        {
            loadedWs.Cells["A1"].PutValue("Modified without password");
            Console.WriteLine("Cell modified successfully (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Modification blocked as expected: " + ex.Message);
        }

        // Cleanup
        wb.Dispose();
        loadedWb.Dispose();
    }
}