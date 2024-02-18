using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System.Linq;

namespace BalanceCompute
{



    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel File (*.xlsx)|*.xlsx"
            };

            var fileResult = openFileDialog1.ShowDialog();

            if (fileResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            string message;

            var rawDatas = LoadRawData(textBox1.Text, out message);

            if (string.IsNullOrEmpty(message) == false)
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + message;
            }

            //讀取銀行交易

            IEnumerable<Translation> transData;

            if (string.IsNullOrEmpty(textBox2.Text))
            {
                transData = new List<Translation>();
            }
            else
            {
                transData = LoadTranslation(textBox2.Text, out message);
            }


            //產生表格
            string fileName = GenerTable(rawDatas, transData, out message);


            if (string.IsNullOrEmpty(message) == false)
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + message;
            }
            else
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + $"完成 路徑: {fileName}";
            }


            
        }


        public static IEnumerable<RawData> LoadRawData(string filePath, out string message)
        {
            message = string.Empty;

            List<RawData> result = new List<RawData>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var lastRow = ws.LastRowUsed().RowNumber();

                int j = 0;

                for (int i = 2; i <= lastRow; i++)
                {


                    RawData temp = new RawData();


                    temp.SerialNo = ws.Cell(i, 4).Value.ToString() ?? string.Empty;


                    string strDate = ws.Cell(i, 8).Value.ToString() ?? string.Empty;

                    if (DateTime.TryParse(strDate, out DateTime date))
                    {
                        temp.Date = date;
                    }
                    else
                    {
                        message = $"第{i}列 日期轉換異常";
                        break;
                    }


                    temp.PayWay = ws.Cell(i, 12).Value.ToString() ?? string.Empty;

                    string strAmount = ws.Cell(i, 15).Value.ToString() ?? string.Empty;

                    if (decimal.TryParse(strAmount, out decimal amount))
                    {
                        temp.Amount = amount;
                    }
                    else
                    {
                        message = $"第{i}列 金額轉換異常";
                        break;
                    }


                    result.Add(temp);
                }
            }

            return result;


        }

        public static IEnumerable<Translation> LoadTranslation(string filePath, out string message)
        {
            message = string.Empty;

            List<Translation> result = new List<Translation>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var lastRow = ws.LastRowUsed().RowNumber();

                int j = 0;

                for (int i = 2; i <= lastRow; i++)
                {
                    //遇空格跳出
                    if (string.IsNullOrEmpty(ws.Cell(i, 1).Value.ToString()))
                    {
                        break;
                    }

                    Translation temp = new Translation();

                    string strDate = ws.Cell(i, 2).Value.ToString() ?? string.Empty;

                    if (DateTime.TryParse(strDate, out DateTime date))
                    {
                        temp.Date = date;
                    }
                    else
                    {
                        message = $"第{i}列 日期轉換異常";
                        break;
                    }


                    string strAmount = ws.Cell(i, 5).Value.ToString() ?? string.Empty;

                    if (decimal.TryParse(strAmount, out decimal amount))
                    {
                        temp.Amount = amount;
                    }
                    else
                    {
                        message = $"第{i}列 金額轉換異常";
                        break;
                    }

                    temp.Remark = ws.Cell(i, 6).Value.ToString() ?? string.Empty;

                    result.Add(temp);
                }
            }

            return result;
        }

        public static string GenerTable(IEnumerable<RawData> rawDatas, IEnumerable<Translation> transDatas, out string message)
        {
            message = string.Empty;


            string[] department = new string[] { "總部大樓", "資策會", "國票" };
            string[] serialNo = new string[] { "00000389", "00000390", "00000505" };

            string[] payment = new string[] { "銀行信用卡", "悠遊卡", "一卡通", "LINE PAY" };
            int[] paymentDay = new int[] { 0, 1, 2, 7 };

            int[] colorR = new int[] { 255, 197, 252, 178 };
            int[] colorG = new int[] { 255, 217, 213, 222 };
            int[] colorB = new int[] { 0, 241, 180, 130 };

            decimal[] feeRate = new decimal[] { 1.8m, 1.5m, 1.5m, 2.2m };

            

            string fileName = AppDomain.CurrentDomain.BaseDirectory + String.Format("{0}-{1}.xlsx"
                , rawDatas.OrderBy(x=>x.Date).FirstOrDefault().Date.ToString("MMdd")
                , rawDatas.OrderByDescending(x => x.Date).FirstOrDefault().Date.ToString("MMdd"));

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("總表");

                int row = 2;

                int col = 1;

                int startCol = 0;

                ws.Column(1).Width = 15;


                for (int i = 0; i < payment.Length; i++)
                {
                    startCol = col + 1;

                    for (int j = 0; j < department.Length; j++)
                    {
                        ws.Cell(row, ++col).SetValue(department[j]);
                        ws.Cell(row, ++col).SetValue($"{feeRate[i]}%");
                    }

                    ws.Cell(row, ++col).SetValue("小計");
                    ws.Cell(row, ++col).SetValue("撥款日");
                    ws.Cell(row, ++col).SetValue("入賬");
                    ws.Range(1, startCol, 1, col).Merge();
                    ws.Cell(1, startCol).SetValue(payment[i]);
                    ws.Cell(1, startCol).Style.Fill.BackgroundColor = XLColor.FromArgb(colorR[i], colorG[i], colorB[i]);
                    ws.Cell(1, startCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }


                

                foreach (var date in rawDatas.GroupBy(x => x.Date).OrderBy(x => x.Key))
                {
                    //日期
                    col = 0;
                    ws.Cell(++row, ++col).SetValue(date.Key.ToString("yyyy/MM/dd"));

                    //付款方式
                    for (int i = 0; i < payment.Length; i++)
                    {
                        decimal total = 0;
                        //部門
                        for (int j = 0; j < department.Length; j++)
                        {
                            var data = rawDatas.Where(x => x.Date == date.Key
                            && x.SerialNo == serialNo[j]
                            && x.PayWay == payment[i]);

                            //金額
                            ws.Cell(row, ++col).SetValue(data.Sum(x => x.Amount));
                            //手續費
                            ws.Cell(row, ++col).SetValue(Math.Round(data.Sum(x => x.Amount) * feeRate[i] * 0.01m, MidpointRounding.AwayFromZero));

                            total += data.Sum(x => x.Amount) - Math.Round(data.Sum(x => x.Amount) * feeRate[i] * 0.01m, MidpointRounding.AwayFromZero);
                        }

                        var cellWithFormulaA1 = ws.Cell(row, ++col);

                        //塞公式
                        switch (i)
                        {
                            case 0:
                                cellWithFormulaA1.FormulaA1 = $"=B{row}-C{row}+D{row}-E{row}+F{row}-G{row}";
                                break;
                            case 1:
                                cellWithFormulaA1.FormulaA1 = $"=K{row}-L{row}+M{row}-N{row}+O{row}-P{row}";
                                break;
                            case 2:
                                cellWithFormulaA1.FormulaA1 = $"=T{row}-U{row}+V{row}-W{row}+X{row}-Y{row}";
                                break;
                            case 3:
                                cellWithFormulaA1.FormulaA1 = $"=AC{row}-AD{row}+AE{row}-AF{row}+AG{row}-AH{row}";
                                break;
                            default:
                                break;
                        }


                        List<Translation> tranData = new List<Translation>();

                        //對帳
                        switch (i)
                        {
                            case 1:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("悠遊卡儲值金")));
                                break;
                            case 2:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("一卡通票證股份有限")));
                                break;
                            case 3:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("國泰世華商業銀行受託信託財產專戶")));
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("一卡通MO提領")));
                                break;
                            default:
                                break;
                        }

                        if(tranData.Count()>0)
                        {
                            ws.Cell(row, ++col).SetValue(tranData.Count() > 0 ? tranData.FirstOrDefault().Date.ToString("MM/dd") : string.Empty);
                            ws.Cell(row, ++col).SetValue(tranData.Sum(x => x.Amount));

                            if(tranData.Sum(x => x.Amount)!= total)
                            {
                                ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Red;
                            }

                        }
                        else
                        {
                            col += 2;
                        }
                    }
                }


                //ws.Columns().AdjustToContents();

                ws.SheetView.Freeze(2, 1);

                wb.SaveAs(fileName);
            }

            return fileName;

        }


        public static IEnumerable<SystemData> LoadSysytemData(string filePath, out DateTime? date, out string message)
        {
            message = string.Empty;

            date = null;

            List<SystemData> data = new List<SystemData>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet("第 1 頁");

                var lastRow = ws.LastRowUsed().RowNumber();

                for (int i = 1; i <= lastRow; i++)
                {
                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    if (rowData.IndexOf("門市代號") > -1)
                    {
                        SystemData temp = new SystemData();

                        var idxSpec = rowData.LastIndexOf(" ");

                        temp.Store = rowData.Substring(idxSpec, rowData.Length - idxSpec).Trim().Replace("圓環", "店");

                        i += 3;

                        decimal.TryParse(ws.Cell(i, 2).Value.ToString() ?? string.Empty, out decimal amount);

                        temp.Cash = amount;

                        if (!date.HasValue)
                        {
                            var dateArray = (ws.Cell(i, 1).Value.ToString() ?? String.Empty).Split("/");

                            date = new DateTime(int.Parse(dateArray[0]) + 1911, int.Parse(dateArray[1]), int.Parse(dateArray[2]));
                        }

                        data.Add(temp);
                    }
                }
            }

            if (data.Count() == 0)
            {
                message = "頂新資料空白";
            }

            return data;
        }

        private static IEnumerable<BalanceData> LoadBalanceData(string filePath, out string title, out string message)
        {
            List<BalanceData> data = new List<BalanceData>();

            message = string.Empty;
            title = string.Empty;

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                var lastRow = ws.LastRowUsed().RowNumber();

                int j = 0;

                while (true)
                {
                    var tempTitle = ws.Cell(1, ++j).Value.ToString() ?? string.Empty;

                    if (string.IsNullOrEmpty(tempTitle))
                    {
                        break;
                    }
                    else
                    {
                        title = tempTitle;
                    }
                }


                for (int i = 2; i <= lastRow; i++)
                {
                    string rowData = ws.Cell(i, 1).Value.ToString() ?? string.Empty;

                    BalanceData temp = new BalanceData();


                    temp.Store = rowData;

                    if (temp.Store.IndexOf("合計") > -1)
                    {
                        break;
                    }

                    decimal.TryParse(ws.Cell(i, 2).Value.ToString() ?? string.Empty, out decimal amount);

                    temp.LastBalance = amount;

                    data.Add(temp);
                }
            }

            return data;
        }

        private static string ExportResult(IEnumerable<BalanceData> _Balance, IEnumerable<SystemData> _System, string title, DateTime date)
        {
            string fileName = AppDomain.CurrentDomain.BaseDirectory + String.Format("{0}.xlsx", date.ToString("MMdd"));

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(date.ToString("MMdd"));

                int i = 1;

                int j = 1;

                ws.Cell(i, j++).SetValue("門市");
                ws.Cell(i, j++).SetValue(title);
                ws.Cell(i, j++).SetValue(string.Format("{0}現金收入", date.ToString("MM/dd")));
                ws.Cell(i, j++).SetValue(string.Format("{0}餘額", date.ToString("MM/dd")));


                foreach (var item in _Balance)
                {
                    var system = _System.FirstOrDefault(x => x.Store.IndexOf(item.Store) > -1);

                    item.Cash = system.Cash;

                    j = 1;

                    ws.Cell(++i, j).SetValue(item.Store);

                    ws.Cell(i, ++j).SetValue(item.LastBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                    ws.Cell(i, ++j).SetValue(item.Cash);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                    ws.Cell(i, ++j).SetValue(item.NowBalance);
                    ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";
                }


                j = 1;
                ws.Cell(++i, j).SetValue("合計");
                ws.Range(i, 1, i, 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x => x.LastBalance));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x => x.Cash));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";

                ws.Cell(i, ++j).SetValue(_Balance.Sum(x => x.NowBalance));
                ws.Cell(i, j).Style.NumberFormat.Format = "#,##0";


                ws.Columns().AdjustToContents();

                wb.SaveAs(fileName);
            }

            return fileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel File (*.xlsx)|*.xlsx"
            };

            var fileResult = openFileDialog1.ShowDialog();

            if (fileResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }
    }
}