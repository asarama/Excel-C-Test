using System.Runtime.InteropServices; // Guid, ClassInterface, ClassInterfaceType
using Excel = Microsoft.Office.Interop.Excel;       //Microsoft Excel 14 object in references-> COM tab

[Guid("97E1D9DB-8478-4E56-9D6D-26D8EF13B100")]
public interface ICSharpTestLibrary
{
    double GetSum(double number_1, double number_2);
}

[Guid("BBF87E31-77E2-46B6-8093-1689A144BFC6")]
[ClassInterface(ClassInterfaceType.None)]
public class CSharpTestLibrary : ICSharpTestLibrary
{
    public double GetSum(double number_1, double number_2)
    {
        return number_1 + number_2;
    }
    public string CurrentRow()
    {

        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"Book.xlsm");
        Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        Excel.Range xlRange = xlWorksheet.UsedRange;

        return "Success";
    }
}