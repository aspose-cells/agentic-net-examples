using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the encrypted workbook using the original password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "oldPassword"; // original password
        Workbook workbook = new Workbook("encrypted_input.ods", loadOptions);

        // Modify some cell values in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Updated Text");
        sheet.Cells["B2"].PutValue(9876);

        // Re‑encrypt the workbook with a new password
        workbook.Settings.Password = "newPassword";

        // Save the workbook as ODS format
        workbook.Save("reencrypted_output.ods", SaveFormat.Ods);
    }
}