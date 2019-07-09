using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DalgakiranImportTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            string x = Program.GetExportTemplate();
            Console.WriteLine(x);
            Console.ReadLine();
        }

        public static string GetExportTemplate()
        {
            string fileName = @"C:\Users\Василий\Desktop\dasdasd.xls";

            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            sheet.Range[1, 1].Text = "ID";
            sheet.Range[1, 2].Text = "Код SAP";
            sheet.Range[1, 3].Text = "Код старый (артикул)";
            sheet.Range[1, 4].Text = "Турецкое название";
            sheet.Range[1, 5].Text = "Английское название";
            sheet.Range[1, 6].Text = "Русское название";
            sheet.Range[1, 7].Text = "Тип";
            sheet.Range[1, 8].Text = "Направление";
            sheet.Range[1, 9].Text = "Вид оборудования";
            sheet.Range[1, 10].Text = "Серия";
            sheet.Range[1, 11].Text = "TR KAR EXWORKS IST, €";
            sheet.Range[1, 12].Text = "2019 Fiyat Listesi, €";
            sheet.Range[1, 13].Text = "Со склада в Стамбуле + 15 %, €";
            sheet.Range[1, 14].Text = "%6 Transport, €";
            sheet.Range[1, 14].Text = "%10 General Expences, €";
            sheet.Range[1, 14].Text = "%25 İskonto Payı, €";
            sheet.Range[1, 14].Text = "%20 НДС, €";

            sheet.Range[1, 1, 1, 14].VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet.Range[1, 1, 1, 14].HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[1, 1, 1, 14].ColumnWidth = 20;
            sheet.Range[1, 1, 1, 14].WrapText = true;
            sheet.Range[1, 1, 1, 14].AutofitRows();

            workbook.SaveAs(fileName);
            workbook.Close();
            excelEngine.Dispose();
            return fileName;
        }

    }
}
