using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DalgakiranImportTemplate
{
    public class ImportTemplate
    {
        public static void Main(string[] args)
        {
            string x = ImportTemplate.GetExportTemplate(false, Guid.Empty, Guid.Empty);
            Console.WriteLine(x);
            Console.ReadLine();
        }

        public static string GetExportTemplate(bool isFirst, Guid categoryId, Guid typeId)
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
            sheet.Range[1, 15].Text = "%10 General Expences, €";
            sheet.Range[1, 16].Text = "15% Profit, €";
            sheet.Range[1, 17].Text = "%25 İskonto Payı, €";
            sheet.Range[1, 18].Text = "%20 НДС, €";

            sheet.Range[1, 1, 1, 18].VerticalAlignment = ExcelVAlign.VAlignCenter;
            sheet.Range[1, 1, 1, 18].HorizontalAlignment = ExcelHAlign.HAlignCenter;
            sheet.Range[1, 1, 1, 18].ColumnWidth = 20;
            sheet.Range[1, 1, 1, 18].WrapText = true;
            sheet.Range[1, 1, 1, 18].AutofitRows();

            if (!isFirst)
            {
                sheet = ImportTemplate.GetExportData(sheet, categoryId, typeId);
            }

            workbook.SaveAs(fileName);
            workbook.Close();
            excelEngine.Dispose();
            return fileName;
        }

        public static IWorksheet GetExportData(IWorksheet sheet, Guid categoryId, Guid typeId)
        {
            Guid cultureId = new Guid("1A778E3F-0A8E-E111-84A3-00155D054C03");
            string sqlText = $@"DECLARE @Culture uniqueidentifier = '{cultureId}';

                                SELECT pr.[Id],
                                       ISNULL(pr.[TcmCodeSap], '') as CodeSap,
	                                   ISNULL(pr.[TcmCodeOld], '') as CodeOld,
	                                   ISNULL(pr.[TcmTurkName], '') as TurkName,
	                                   ISNULL(pr.[TcmEngName], '') as EngName,
	                                   ISNULL(pr.[Name], '') as RuName,
	                                   ISNULL(prt.[Name], '') as [Type],
	                                   ISNULL(prd.[Name], '') as Direction,
	                                   ISNULL(prte.[Name], '') as TypeEquipment,
	                                   ISNULL(prs.[Name], '') as Serial,
	                                   pr.[TcmTRKarExworksEuro],
	                                   pr.[TcmFiyatListesiEuro],
	                                   pr.[TcmFromStockToIstanbulEuro],
	                                   pr.[TcmTransportEuro],
	                                   pr.[TcmGeneralExpencesEuro],
	                                   pr.[TcmProfitEuro],
	                                   pr.[TcmIskontoPayEuro],
	                                   pr.[TcmNDSEuro]
                                  FROM [Product] pr
                                       LEFT OUTER JOIN [SysProductCategoryLcz] prt on prt.[RecordId] = pr.[CategoryId] and prt.[SysCultureId] = @Culture
	                                   LEFT OUTER JOIN [SysProductTypeLcz] prd on prd.[RecordId] = pr.[TypeId] and prd.[SysCultureId] = @Culture
	                                   LEFT OUTER JOIN [SysTcmTypeEquipmentLcz] prte on prte.[RecordId] = pr.[TcmTypeEquipmentId] and prte.[SysCultureId] = @Culture
	                                   LEFT OUTER JOIN [SysTcmProductSerialLcz] prs on prs.[RecordId] = pr.[TcmProductSerialId] and prs.[SysCultureId] = @Culture";

            if (categoryId != Guid.Empty || typeId != Guid.Empty)
            {
                sqlText += " WHERE ";
                if (categoryId != Guid.Empty)
                {
                    sqlText += $@"pr.[CategoryId] = '{categoryId}'";

                    if (typeId != Guid.Empty)
                    {
                        sqlText += $@" AND ";
                    }
                }
                if (typeId != Guid.Empty)
                {
                    sqlText += $@"pr.[TypeId] = '{typeId}'";
                }
            }

            DataTable dt = ImportTemplate.GetDataTable(sqlText);

            int r = 2;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sheet.Range[r, 1].Text = dt.Rows[i].Field<Guid>("Id").ToString();
                sheet.Range[r, 2].Text = dt.Rows[i].Field<string>("CodeSap");
                sheet.Range[r, 3].Text = dt.Rows[i].Field<string>("CodeOld");
                sheet.Range[r, 4].Text = dt.Rows[i].Field<string>("TurkName");
                sheet.Range[r, 5].Text = dt.Rows[i].Field<string>("EngName");
                sheet.Range[r, 6].Text = dt.Rows[i].Field<string>("RuName");
                sheet.Range[r, 7].Text = dt.Rows[i].Field<string>("Type");
                sheet.Range[r, 8].Text = dt.Rows[i].Field<string>("Direction");
                sheet.Range[r, 9].Text = dt.Rows[i].Field<string>("TypeEquipment");
                sheet.Range[r, 10].Text = dt.Rows[i].Field<string>("Serial");
                sheet.Range[r, 11].Number = dt.Rows[i].Field<int>("TcmTRKarExworksEuro");
                sheet.Range[r, 12].Number = dt.Rows[i].Field<int>("TcmFiyatListesiEuro");
                sheet.Range[r, 13].Number = dt.Rows[i].Field<int>("TcmFromStockToIstanbulEuro");
                sheet.Range[r, 14].Number = dt.Rows[i].Field<int>("TcmTransportEuro");
                sheet.Range[r, 15].Number = dt.Rows[i].Field<int>("TcmGeneralExpencesEuro");
                sheet.Range[r, 16].Number = dt.Rows[i].Field<int>("TcmProfitEuro");
                sheet.Range[r, 17].Number = dt.Rows[i].Field<int>("TcmIskontoPayEuro");
                sheet.Range[r, 18].Number = dt.Rows[i].Field<int>("TcmNDSEuro");
                r++;
            }
            sheet.Range[2, 1, r, 10].AutofitColumns();

            sheet.Range[r, 1].Text = "123123";
            sheet.Range[r, 2].Text = "123123";
            sheet.Range[r, 3].Text = "123123";
            sheet.Range[r, 4].Text = "123123";
            sheet.Range[r, 5].Text = "123123";
            sheet.Range[r, 6].Text = "123123";
            sheet.Range[r, 7].Text = "123123";
            sheet.Range[r, 8].Text = "123123";
            sheet.Range[r, 9].Text = "123123";
            sheet.Range[r, 10].Text = "123123";
            sheet.Range[r, 11].Number = 1;
            sheet.Range[r, 12].Number = 2;
            sheet.Range[r, 13].Number = 3;
            sheet.Range[r, 14].Number = 5;
            sheet.Range[r, 15].Number = 6;
            sheet.Range[r, 16].Number = 7;
            sheet.Range[r, 17].Number = 8;
            sheet.Range[r, 18].Number = 9;

            return sheet;
        }

        public static DataTable GetDataTable(string sqlText)
        {
            SqlConnection connection = new SqlConnection(@"Data Source = 172.16.2.75; Initial Catalog = Dalgakiran7141; Persist Security Info = True; MultipleActiveResultSets = True; Integrated Security = False; User ID = SA; Password = !QAZxsw2@WSXzaq1; Pooling = true; Max Pool Size = 100; Async = true; Connection Timeout = 500");

            connection.Open();

            SqlCommand command = new SqlCommand(sqlText, connection);
            SqlDataReader reader = command.ExecuteReader();
            DataTable dt = new DataTable();

            dt.Load(reader);
            connection.Close();

            return dt;
        }
    }
}
