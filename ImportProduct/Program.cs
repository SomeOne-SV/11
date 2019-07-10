using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DalgakiranImportTemplate
{
    class ImportProduct
    {
        static void Main(string[] args)
        {
            string x = ImportTemplate.GetExportTemplate(false, Guid.Empty, Guid.Empty);

            using (Stream fileStream = File.OpenRead(x))
            {
                MemoryStream stream = new MemoryStream();
                fileStream.CopyTo(stream);
                stream.Seek(0, SeekOrigin.Begin);
                using (ExcelEngine engine = new ExcelEngine())
                {
                    IWorkbook workbook = engine.Excel.Workbooks.Open(stream);
                    IWorksheet sheet = workbook.Worksheets[0];
                    if (!ImportProduct.CheckFile(sheet))
                    {
                        sheet.Range[1, 19].Text = @"Шаблон отличается";
                    }
                    else
                    {
                        for (int j = 1; j < sheet.Rows.Length; j++)
                        {
                            ImportProduct.ParseProduct(sheet.Rows[j], sheet);
                        }
                    }
                    sheet.Range[1, 19].AutofitRows();
                    workbook.SaveAs(@"C:\Users\Василий\Desktop\dasdasd2.xls");
                    workbook.Close();
                    engine.Dispose();
                }

            }

        }
        
        public static void ParseProduct(IRange row, IWorksheet sheet)
        {
            try
            {
                Product product = new Product();
                Guid.TryParse(row.Cells[0].Text, out Guid resultId);
                product.Id = resultId;
                product.CodeSap = row.Cells[1].Text;
                product.CodeOld = row.Cells[2].Text;
                product.TurkName = row.Cells[3].Text;
                product.EngName = row.Cells[4].Text;
                product.RuName = row.Cells[5].Text;
                product.Type = GetLookupValue(row.Cells[6].Text, @"SysProductCategoryLcz");
                product.Direction = GetLookupValue(row.Cells[7].Text, @"SysProductTypeLcz");
                product.TypeEquipment = GetLookupValue(row.Cells[8].Text, @"SysTcmTypeEquipmentLcz");
                product.Serial = GetLookupValue(row.Cells[9].Text, @"SysTcmProductSerialLcz");
                product.TcmTRKarExworksEuro = GetIntValue(row.Cells[10]);
                product.TcmFiyatListesiEuro = GetIntValue(row.Cells[11]);
                product.TcmFromStockToIstanbulEuro = GetIntValue(row.Cells[12]);
                product.TcmTransportEuro = GetIntValue(row.Cells[13]);
                product.TcmGeneralExpencesEuro = GetIntValue(row.Cells[14]);
                product.TcmProfitEuro = GetIntValue(row.Cells[15]);
                product.TcmIskontoPayEuro = GetIntValue(row.Cells[16]);
                product.TcmNDSEuro = GetIntValue(row.Cells[17]);

                CheckRequearedFields(product, row, sheet);
            }
            catch(Exception ex)
            {
                sheet.Range[row.Row, 1, row.Row, 18].CellStyle.Color = System.Drawing.Color.Red;
                sheet.Range[row.Row, 19].Text = ex.Message;
            }
        }

        public static void CheckRequearedFields(Product product, IRange row, IWorksheet sheet)
        {
            if (product.Type == Guid.Empty || product.Direction == Guid.Empty ||
                product.Serial == Guid.Empty || product.CodeSap == String.Empty)
            {
                sheet.Range[row.Row, 1, row.Row, 18].CellStyle.Color = System.Drawing.Color.Red;
                sheet.Range[row.Row, 19].Text = @"Не заполнены обязательные поля";
            }
            else
            {
                CheckInBpm(product, row, sheet);
            }
        }

        public static void CheckInBpm(Product product, IRange row, IWorksheet sheet)
        {
            Guid result = Guid.Empty;

            string sqlText = $@"SELECT TOP 1 [Id]
                                  FROM [Product]
                                 WHERE [Id] = '{product.Id}' ";

            if (product.CodeSap != null || product.CodeSap != String.Empty)
            {
                sqlText += $@" OR [TcmCodeSap] = N'{product.CodeSap}'";
            }

            if (product.CodeOld != null || product.CodeOld != String.Empty)
            {
                sqlText += $@" OR [TcmCodeOld] = N'{product.CodeOld}' ";
            }
            DataTable dt = ImportProduct.GetDataTable(sqlText);

            if (dt.Rows.Count > 0)
            {
                sheet.Range[row.Row, 1, row.Row, 18].CellStyle.Color = System.Drawing.Color.Red;
                sheet.Range[row.Row, 19].Text = @"Продукт уже существует";
            }
            else
            {
                //InsertProductInBPM(product, sheet);
            }

        }

        public static int GetIntValue(IRange cell)
        {
            int result = 0;
            if (cell.HasNumber)
            {
                result = Convert.ToInt32(cell.Number);
            }
            else
            {
                result = Convert.ToInt32(cell.Text);
            }
            return result;
        }

        public static Guid GetLookupValue(string name, string lookupName)
        {
            Guid result = Guid.Empty;

            string sqlText = $@"SELECT TOP 1 [RecordId]
                                  FROM [{lookupName}]
                                 WHERE [Name] = N'{name}'";

            DataTable dt = ImportProduct.GetDataTable(sqlText);

            if (dt.Rows.Count > 0)
            {
                result = dt.Rows[0].Field<Guid>("RecordId");
            }
            return result;

        }

        public static bool CheckFile(IWorksheet sheet)
        {
            if (sheet.Range[1, 1].Text == "ID" && 
                sheet.Range[1, 2].Text == "Код SAP" && 
                sheet.Range[1, 3].Text == "Код старый (артикул)" && 
                sheet.Range[1, 4].Text == "Турецкое название" && 
                sheet.Range[1, 5].Text == "Английское название" && 
                sheet.Range[1, 6].Text == "Русское название" && 
                sheet.Range[1, 7].Text == "Тип" && 
                sheet.Range[1, 8].Text == "Направление" && 
                sheet.Range[1, 9].Text == "Вид оборудования" && 
                sheet.Range[1, 10].Text == "Серия" && 
                sheet.Range[1, 11].Text == "TR KAR EXWORKS IST, €" && 
                sheet.Range[1, 12].Text == "2019 Fiyat Listesi, €" && 
                sheet.Range[1, 13].Text == "Со склада в Стамбуле + 15 %, €" && 
                sheet.Range[1, 14].Text == "%6 Transport, €" && 
                sheet.Range[1, 15].Text == "%10 General Expences, €" && 
                sheet.Range[1, 16].Text == "15% Profit, €" && 
                sheet.Range[1, 17].Text == "%25 İskonto Payı, €" && 
                sheet.Range[1, 18].Text == "%20 НДС, €")
            {
                return true;
            }
            return false;
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

    public class Product
    {
        public Guid Id { get; set; }

        public string CodeSap { get; set; }

        public string CodeOld { get; set; }

        public string TurkName { get; set; }

        public string EngName { get; set; }

        public string RuName { get; set; }

        public Guid Type { get; set; }

        public Guid Direction { get; set; }

        public Guid TypeEquipment { get; set; }

        public Guid Serial { get; set; }

        public int TcmTRKarExworksEuro { get; set; }

        public int TcmFiyatListesiEuro { get; set; }

        public int TcmFromStockToIstanbulEuro { get; set; }

        public int TcmTransportEuro { get; set; }

        public int TcmGeneralExpencesEuro { get; set; }

        public int TcmProfitEuro { get; set; }

        public int TcmIskontoPayEuro { get; set; }

        public int TcmNDSEuro { get; set; }
    }
}
