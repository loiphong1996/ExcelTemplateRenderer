using OfficeOpenXml;

namespace ExcelLibs.ExcelElement
{
    public interface IExcelElement
    {
        int SetupSpace(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset);
        int ApplyValue(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset);
        void ApplyStyle(ExcelWorksheet templateSheet, ExcelWorksheet outputSheet, int rowOffset);
    }
}