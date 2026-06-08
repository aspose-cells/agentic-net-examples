using System;
using System.Text;
using Aspose.Cells;

class EmbedCustomXmlProtectionPolicy
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "StrongPassword123");

        // Define a custom XML part that describes the protection policy
        string xmlData = @"<ProtectionPolicy>
                               <AllowEditLockedCells>false</AllowEditLockedCells>
                               <AllowFormattingCells>true</AllowFormattingCells>
                           </ProtectionPolicy>";

        // Schema can be empty if not required
        string xmlSchema = string.Empty;

        // Add the custom XML part to the workbook
        workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlData), Encoding.UTF8.GetBytes(xmlSchema));

        // Save the protected workbook with the embedded custom XML part
        workbook.Save("ProtectedWorkbookWithPolicy.xlsx");
    }
}