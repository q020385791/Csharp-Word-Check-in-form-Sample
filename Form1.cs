using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Nmaereport2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Object oMissing = System.Reflection.Missing.Value;
            Object oEndOfDoc = "\\endofdoc";
            Microsoft.Office.Interop.Word.Application wrdApp;

            Microsoft.Office.Interop.Word._Document myDoc;
            wrdApp = new Microsoft.Office.Interop.Word.Application();
            //執行過程不在畫面上開啟 Word
            wrdApp.Visible = false;    
            myDoc = wrdApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //邊界設定
            //上邊界
            wrdApp.ActiveDocument.PageSetup.TopMargin = 23F;

            //左邊界
            wrdApp.ActiveDocument.PageSetup.LeftMargin = 20F;

            //下邊界
            wrdApp.ActiveDocument.PageSetup.BottomMargin = 10F;

            //右邊界
            wrdApp.ActiveDocument.PageSetup.RightMargin = 20F;

            //加入文字
            Microsoft.Office.Interop.Word.Paragraph oPara;
            object oRng = myDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = myDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = "附件6";

            //設定粗體
            oPara.Range.Font.Bold = 1;
            //oPara.Format.SpaceAfter = 6;

            //靠右對齊
            oPara.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            oPara.Range.InsertParagraphAfter();

            wrdApp.Application.Selection.MoveEnd();
            object count = 14;
            //换一行;
            object WdLine = Word.WdUnits.wdLine;
            wrdApp.Selection.MoveDown(ref WdLine, ref count, ref oMissing);//移动焦点
            //wrdApp.Selection.TypeParagraph();//插入段落
            oPara.Range.Font.Size = 16;
            oPara.Range.Font.Bold = 0;

            //Selected Year and Month
            int Year = (int)(Convert.ToInt64(cbYear.Text));
            int month = (int)Convert.ToInt64(cbMonth.Text);


            //靠左對齊
            oPara.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            oPara.Range.Text = (Year - 1911).ToString() + "年" + month + "月份              廠商駐點簽到表";
            oPara.Range.InsertParagraphAfter();

            //移动焦点
            wrdApp.Selection.MoveDown(ref WdLine, ref count, ref oMissing);

            //字體大小
            oPara.Range.Font.Size = 12;

            //字體樣式
            oPara.Range.Font.Name = "標楷體";

            //創建Table
            //Create Table
            //Create a Table with one row 、 six Column
            Table table = oPara.Range.Tables.Add(wrdApp.Selection.Range, 1, 6, oMissing, ref oMissing);

            //每一行的高度
            table.Rows.Height = 25f;

            //內框線樣式
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //外框線樣式
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            //外框線粗度
            table.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth150pt;

            //First Line(Title)
            //表格第一行(標題)
            table.Cell(1, 1).Range.Text = "日期";
            //All Line Fill Gray Color
            //整行加入灰色
            table.Cell(1, 1).Row.Shading.BackgroundPatternColor = WdColor.wdColorGray125;
            table.Cell(1, 2).Range.Text = "廠商人員簽到";
            table.Cell(1, 3).Range.Text = "到院時間";
            table.Cell(1, 4).Range.Text = "離院時間";
            table.Cell(1, 5).Range.Text = "資訊部門\n代表簽章";
            table.Cell(1, 6).Range.Text = "備註";

            //紀錄已經記錄到陣列的第幾位
            int CountLocationday = 0;
            //紀錄星期五在第幾行 塞到陣列去
            int[] LocationFriday = new int[] { 0, 0, 0, 0, 0 };
            for (int i = 0; i < DateTime.DaysInMonth(Year, month); i++)
            {
                //Get the DataTime You Selected
                //取得所選的時間基準年月日
                DateTime dt = new DateTime(Year, month, i + 1);


                //Only Add row between Monday and Friday 
                //如果不是星期六或是星期日才增加一行欄位
                if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                {
                    var row = table.Rows.Add();

                    //背景設定回白色，因為第一行原本背景為白色
                    row.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
                    row.Range.Text = month + "/" + (i + 1).ToString();
                    if (dt.DayOfWeek == DayOfWeek.Friday)
                    {
                        int CountFriday = row.Index;
                        LocationFriday.SetValue(CountFriday, CountLocationday);
                        CountLocationday++;

                        //It's not working,but I can't find reason,so I use other way
                        //這要畫框線沒用只有對最後一個有用 不知道為啥
                        //var TestIndex = row.Index; tt
                        //table.Cell(TestIndex, i).Range.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
                    }
                }
            }

            //上面挑選完星期五在哪邊，就可以開始畫粗黑框了
            //注意 寫在上面if (dt.DayOfWeek == DayOfWeek.Friday)那邊是沒有用的，
            foreach (var Friday in LocationFriday)
            {
                if (Friday != 0)
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        table.Cell(Friday, i).Range.Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
                    }
                }
            }
                string dummyFileName = "SaveHere";

                SaveFileDialog sf = new SaveFileDialog();
                sf.FileName = dummyFileName;
                if (sf.ShowDialog() == DialogResult.OK)
                {
                    string savePath = Path.GetDirectoryName(sf.FileName);
                    //另存文件
                    Object oSavePath = savePath+"\\"+ dummyFileName;    //存檔路徑
                    Object oFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;    //格式
                    myDoc.SaveAs(ref oSavePath, ref oFormat,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                ref oMissing, ref oMissing
                            );
                    //關閉檔案
                    Object oFalse = false;
                    myDoc.Close(ref oFalse, ref oMissing, ref oMissing);
                    Console.WriteLine("創建完畢");
                }

                

            




        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //年份初始值
            for (int i = DateTime.Now.Year - 5; i <= DateTime.Now.Year + 5; i++)
            {
                cbYear.Items.Add(i);
            }
            cbYear.Text = DateTime.Now.Year.ToString();
            //月份初始值
            for (int i = 1; i <= 12; i++)
            {
                cbMonth.Items.Add(i);
            }
            cbMonth.Text = DateTime.Now.Month.ToString();



        }
    }
}
