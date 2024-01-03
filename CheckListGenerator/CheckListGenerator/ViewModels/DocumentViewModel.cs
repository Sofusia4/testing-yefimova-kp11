using ClosedXML.Excel;

namespace CheckListGenerator.ViewModels
{
    public class DocumentViewModel
    {
        public XLWorkbook Workbook { get; set; }
        public string Path { get; set; }
        public string FileName { get; set; }
    }
}
