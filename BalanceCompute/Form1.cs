using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System.Linq;
using System.Windows.Forms;

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

            //Ū���Ȧ���

            IEnumerable<Translation> transData;

            if (string.IsNullOrEmpty(textBox2.Text))
            {
                transData = new List<Translation>();
            }
            else
            {
                transData = LoadTranslation(textBox2.Text, out message);
            }


            //���ͪ��
            string fileName = GenerTable(rawDatas, transData, out message);


            if (string.IsNullOrEmpty(message) == false)
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + message;
            }
            else
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + $"���� ���|: {fileName}";
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
                        message = $"��{i}�C ����ഫ���`";
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
                        message = $"��{i}�C ���B�ഫ���`";
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
                    //�J�Ů���X
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
                        message = $"��{i}�C ����ഫ���`";
                        break;
                    }


                    string strAmount = ws.Cell(i, 5).Value.ToString() ?? string.Empty;

                    if (decimal.TryParse(strAmount, out decimal amount))
                    {
                        temp.Amount = amount;
                    }
                    else
                    {
                        message = $"��{i}�C ���B�ഫ���`";
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


            string[] department = new string[] { "�`���j��", "�굦�|", "�겼" };
            string[] serialNo = new string[] { "00000389", "00000390", "00000505" };

            string[] payment = new string[] { "�Ȧ�H�Υd", "�y�C�d", "�@�d�q", "LINE PAY" };
            int[] paymentDay = new int[] { 0, 1, 2, 7 };

            int[] colorR = new int[] { 255, 197, 252, 178 };
            int[] colorG = new int[] { 255, 217, 213, 222 };
            int[] colorB = new int[] { 0, 241, 180, 130 };

            decimal[] feeRate = new decimal[] { 1.8m, 1.5m, 1.5m, 2.2m };



            string fileName = AppDomain.CurrentDomain.BaseDirectory + String.Format("{0}-{1}.xlsx"
                , rawDatas.OrderBy(x => x.Date).FirstOrDefault().Date.ToString("MMdd")
                , rawDatas.OrderByDescending(x => x.Date).FirstOrDefault().Date.ToString("MMdd"));

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("�`��");

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

                    ws.Cell(row, ++col).SetValue("�p�p");
                    ws.Cell(row, ++col).SetValue("���ڤ�");
                    ws.Cell(row, ++col).SetValue("�J��");
                    ws.Range(1, startCol, 1, col).Merge();
                    ws.Cell(1, startCol).SetValue(payment[i]);
                    ws.Cell(1, startCol).Style.Fill.BackgroundColor = XLColor.FromArgb(colorR[i], colorG[i], colorB[i]);
                    ws.Cell(1, startCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }




                foreach (var date in rawDatas.GroupBy(x => x.Date).OrderBy(x => x.Key))
                {
                    //���
                    col = 0;
                    ws.Cell(++row, ++col).SetValue(date.Key.ToString("yyyy/MM/dd"));

                    //�I�ڤ覡
                    for (int i = 0; i < payment.Length; i++)
                    {
                        decimal total = 0;
                        //����
                        for (int j = 0; j < department.Length; j++)
                        {
                            var data = rawDatas.Where(x => x.Date == date.Key
                            && x.SerialNo == serialNo[j]
                            && x.PayWay == payment[i]);

                            //���B
                            ws.Cell(row, ++col).SetValue(data.Sum(x => x.Amount));
                            //����O
                            ws.Cell(row, ++col).SetValue(Math.Round(data.Sum(x => x.Amount) * feeRate[i] * 0.01m, MidpointRounding.AwayFromZero));

                            total += data.Sum(x => x.Amount) - Math.Round(data.Sum(x => x.Amount) * feeRate[i] * 0.01m, MidpointRounding.AwayFromZero);
                        }

                        var cellWithFormulaA1 = ws.Cell(row, ++col);

                        //�뤽��
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

                        //��b
                        switch (i)
                        {
                            case 1:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("�y�C�d�x�Ȫ�")));
                                break;
                            case 2:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("�@�d�q���Ҫѥ�����")));
                                break;
                            case 3:
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("����@�ذӷ~�Ȧ���U�H�U�]���M��")));
                                tranData.AddRange(transDatas.Where(x => x.Date == date.Key.AddDays(paymentDay[i]) && x.Remark.Contains("�@�d�qMO����")));
                                break;
                            default:
                                break;
                        }

                        if (tranData.Count() > 0)
                        {
                            ws.Cell(row, ++col).SetValue(tranData.Count() > 0 ? tranData.FirstOrDefault().Date.ToString("MM/dd") : string.Empty);
                            ws.Cell(row, ++col).SetValue(tranData.Sum(x => x.Amount));

                            if (tranData.Sum(x => x.Amount) != total)
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

                ws.SheetView.Freeze(2, 1);

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

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                DefaultExt = "xlsx",
                Filter = "Excel File (*.xlsx)|*.xlsx"
            };

            var fileResult = openFileDialog1.ShowDialog();

            if (fileResult == System.Windows.Forms.DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            string message = string.Empty;

            var path = GenerSystemFile(out message);


            if (string.IsNullOrEmpty(message) == false)
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + message;
            }
            else
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + $"���� ���|: {path}";
            }


        }


        public static string GenerSystemFile(out string message)
        {
            message = string.Empty;


            string[] department = new string[] { "�`���j��", "�굦�|", "�겼" };
            string[] serialNo = new string[] { "00000389", "00000390", "00000505" };

            string[] payment = new string[] { "�Ȧ�H�Υd", "�y�C�d", "�@�d�q", "LINE PAY" };
            int[] paymentDay = new int[] { 0, 1, 2, 7 };

            int[] colorR = new int[] { 255, 197, 252, 178 };
            int[] colorG = new int[] { 255, 217, 213, 222 };
            int[] colorB = new int[] { 0, 241, 180, 130 };


            string[] title1 = new string[] { "BUKRS", "BLART", "BLDAT", "BUDAT", "MONAT", "BKTXT", "WAERS", "LDGRP", "KURSF_EXT", "WWERT", "XBLNR", "PARGB_HDR", "XMWST" };
            string[] title2 = new string[] { "*���q�N�X (4)", "*��O�b�������� (2)", "*��O�b�������", "*�L�b���", "�|�p���� (2)", "�����Y���� (25)", "*������O (5)", "�����b�s�� (4)", "�ײv (12)", "���O������", "�ѦҤ�󸹽X (16)", "�٦�~�Ƚd�� (4)", "�۰ʭp��|�� (1)" };

            string[] detailTitel1 = new string[] { "BUKRS", "HKONT", "SGTXT", "WRSOL", "WRHAB", "DMBTR", "DMBE2", "MWSKZ", "TXJCD", "KOSTL", "PRCTR", "AUFNR", "PS_POSID", "VALUT", "HBKID", "HKTID", "ZUONR", "VBUND", "SEGMENT" };
            string[] detailTitel2 = new string[] { "���q�N�X (4)", "�`�b��� (10)", "���ؤ��� (50)", "�ɤ�", "�U��", "���B�q���q�N�X���O�r", "Amount in second local currency (LC2)"
                , "�|�X (2)", "���|�����v (15)", "�������� (10)", "�Q���� (10)", "�q�渹�X (12)", "WBS ���� (24)", "�_����", "���ӻȦ� (5)", "���ӻȦ�b�� (5)","�������X (18)","�T���٦� (6)","�Ϭq����s�@���Ϭq (10)" };



            decimal[] feeRate = new decimal[] { 1.8m, 1.5m, 1.5m, 2.2m };


            //string dateFormat = String.Format("{0}-{1}"
            //    , rawDatas.OrderBy(x => x.Date).FirstOrDefault().Date.ToString("MMdd")
            //    , rawDatas.OrderByDescending(x => x.Date).FirstOrDefault().Date.ToString("MMdd"));

            string dateFormat = "TEST";

            string fileName = AppDomain.CurrentDomain.BaseDirectory + $"����{dateFormat}.xlsx";

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(dateFormat);

                int row = 0;

                int col = 0;


                ws.Cell(++row, ++col).SetValue("�W�Ǥ@���O�b����");
                ws.Cell(++row, col).SetValue("// To add field columns to the template, please add technical names.");
                ws.Cell(row, col).Style.Alignment.WrapText = false;
                ws.Cell(++row, col).SetValue("// For a complete list of field columns and their technical names, choose ?in the right upper corner of the app screen and then view the Browseentry of web assistance.\r\n");
                ws.Cell(row, col).Style.Alignment.WrapText = false;              
                ws.Cell(++row, col).SetValue("�妸 ID");
                ws.Range(row, col, row, col + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 192, 0);
                

                for (int i = 0; i < 2; i++)
                {
                    ws.Cell(row += 3, 1).SetValue(i+1);
                    ws.Cell(row , 2).SetValue("���Y");
                    ws.Range(row, 1, row, 20).Style.Fill.BackgroundColor = XLColor.FromArgb(142, 169, 219);

                    col = 1;
                    //8

                    ++row;
                    col = 1;
                    for (int k = 0; k < title1.Length; k++)
                    {
                        ws.Cell(row, ++col).SetValue(title1[k]);

                        ws.Cell(row+1, col).SetValue(title2[k]);

                    }

                    ws.Range(row, 2, row, title1.Length+1).Style.Fill.BackgroundColor = XLColor.FromArgb(237, 241, 249);
                    ws.Range(row+1, 2, row+1, title1.Length + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(237, 241, 249);


                    //��ƦC
                    row+=2;

                    //�Ť@��
                    ++row;
                    ws.Cell(++row, 2).SetValue("���Ӷ���");
                    ws.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.FromArgb(142, 169, 219);

                    ++row;
                    ws.Range(row, 5, row, 6).Merge();
                    ws.Cell(row, 5).SetValue("������O");
                    ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromArgb(237, 241, 249);

                    row++;
                    col = 1;
                    for (int k = 0; k < detailTitel1.Length; k++)
                    {
                        ws.Cell(row, ++col).SetValue(detailTitel1[k]);

                        ws.Cell(row + 1, col).SetValue(detailTitel2[k]);
                    }

                    ws.Range(row, 2, row, detailTitel1.Length + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(237, 241, 249);
                    ws.Range(row + 1, 2, row + 1, detailTitel1.Length + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(237, 241, 249);

                    //��ƦC

                }                

                wb.SaveAs(fileName);
            }

            return fileName;

        }

    }
}