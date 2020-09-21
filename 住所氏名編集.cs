//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;


namespace 口酒井農業水利組合郵送会員住所録
{
    public partial class 住所氏名編集Form : Form
    {
        private string[] Values;
        public 住所録リストForm formMain;
        //public string path = (@"D:\develope\口酒井\住所録\");
        public string 水利関係住所録WB = (@"C:\dropbox\住所録\水利関係住所録.xlsx");


        //住所録呼出
        //dynamic xlApp;
        //dynamic xlBooks;
        //dynamic xlBook;

        public 住所氏名編集Form()
        {
            InitializeComponent();
            分類セット();
            ValuesAttach();

            ////住所録呼出
            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(水利関係住所録WB);
        }


        private void 分類セット()
        {
            分類combo.Items.Add("入作関係");
            分類combo.Items.Add("企業協力金");
            分類combo.Items.Add("墓地管理");
            分類combo.Text = "入作関係";
        }

        private void ValuesAttach()
        {
            Values = new string[5];
            Values[0] = IDBox.Text;
            Values[1] = 氏名Box.Text;
            Values[2] = 郵便番号Box.Text;
            Values[3] = 住所Box.Text;
            Values[4] = 分類combo.Text;
        }


        public string[]  AddSet
        {
            set
            {
                Values = value;
                IDBox.Text = Values[0];
                氏名Box.Text = Values[1];
                郵便番号Box.Text = Values[2];
                住所Box.Text = Values[3];
                分類combo.Text = Values[4];
            }
            get
            {
                return Values;
            }

        }



        private void 住所氏名編集_Load(object sender, EventArgs e)
        {
            ////住所録呼出
            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(@"D:\develope\口酒井\住所録\水利関係住所録.xlsx");

            //// シートを選択
            //var sashidasinin = xlBook.Sheets["差出人"];

            //// セルの領域を選択
            //var 住所氏名 = sashidasinin.Range["A2", "B2"];

            //// 選択した領域の値をメモリー上に格納
            //object[,] Sender = 住所氏名.Value;

            ////MessageBox.Show((string)Sender[1, 1]);
            ////MessageBox.Show((string)Sender[1, 2]);

            //// 選択した領域の値をメモリー上に格納
            //住所Box.Text = (string)Sender[1, 1];
            //郵便番号Box.Text = (string)Sender[1, 2];
            //氏名Box.Text = (string)Sender[1, 2];

            //// 使用したCOMオブジェクトを解放
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(住所氏名);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sashidasinin);

            //// Excelのクローズ
            //xlBook.Close();
            //xlApp.Quit();

            ////// 使用したCOMオブジェクトを解放その２
            ////System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlBook);
            ////System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlBooks);
            ////System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }

        private void 新規追加btn_Click(object sender, EventArgs e)
        {
            if (分類combo.SelectedIndex == -1)
            {
                MessageBox.Show("分類が選択されていません。");
                return;
            }

            //住所録呼出
            //dynamic xlApp = new Excel.Application();
            dynamic xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlBooks = xlApp.Workbooks;
            dynamic xlBook = xlBooks.Open(水利関係住所録WB);

            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(水利関係住所録WB);

            ValuesAttach();

            // シートを選択
            var 口酒井名簿 = xlBook.Sheets["口酒井名簿"];

            //// セルの領域を選択
            var TableRange = 口酒井名簿.Range["A2"].CurrentRegion;
            int i = TableRange.Rows.Count;

            //// 選択した領域の値をメモリー上に格納
            //object[,] values = TableRange.Value;
            string a = 口酒井名簿.range["A" + i].text;
            int b = int.Parse(a);
            int c = b + 1;

            //MessageBox.Show("TableRange.Rows.Count = " + i + ",\n"
            //    + " 口酒井名簿.range[\"A\" + i].text = " + a + ",\n"
            //    + "口酒井名簿.range[\"A\" + i].text のパース = " + b + ",\n"
            //    + "b + 1 = " + (b + 1) + ",\n"
            //    + "c.ToString() = " + (c.ToString()));

            口酒井名簿.range["A" + (i + 1)] = c.ToString();          //ID
            口酒井名簿.range["B" + (i + 1)] = (String)Values[1];   //氏名
            口酒井名簿.range["C" + (i + 1)] = (String)Values[2];   //郵便番号
            口酒井名簿.range["D" + (i + 1)] = (String)Values[3];   //住所
            口酒井名簿.range["E" + (i + 1)] = (String)Values[4];   //分類

            IDBox.Text = c.ToString();          //IDを変更
            ValuesAttach();

            xlBook.Save();
 
            MessageBox.Show("追加しました");

            // 使用したCOMオブジェクトを解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(口酒井名簿);

            // Excelのクローズ
            xlBook.Close();
            xlBooks.Close();
            xlApp.Quit();
            //// 使用したCOMオブジェクトを解放その２
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

            formMain.Lvflag = "追加";

            this.Close();
        }


        private void 修正btn_Click(object sender, EventArgs e)
        {

            ValuesAttach();

            //住所録呼出
            //dynamic xlApp = new Excel.Application();
            dynamic xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlBooks = xlApp.Workbooks;
            dynamic xlBook = xlBooks.Open(水利関係住所録WB);

            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(水利関係住所録WB);


            // シートを選択
            var 口酒井名簿 = xlBook.Sheets["口酒井名簿"];

            //// セルの領域を選択
            var TableRange = 口酒井名簿.Range["A2"].CurrentRegion;

            // 選択した領域の値をメモリー上に格納
            object[,] values = TableRange.Value;


            for (int i = 2; i <= values.GetLength(0); i++)
            {
                double セル値 = (double)values[i, 1];

                if (セル値 == double.Parse(Values[0]) )
                {
                    口酒井名簿.range["B" + i] = (String)Values[1];   //氏名
                    口酒井名簿.range["C" + i] = (String)Values[2];   //郵便番号
                    口酒井名簿.range["D" + i] = (String)Values[3];   //住所
                    口酒井名簿.range["E" + i] = (String)Values[4];   //分類

                    xlBook.Save();
                    MessageBox.Show("修正しました");
                    break;
                }
            }

            // 使用したCOMオブジェクトを解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(口酒井名簿);
            // Excelのクローズ
            xlBook.Close();
            xlBooks.Close();
            xlApp.Quit();
            //// 使用したCOMオブジェクトを解放その２
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

            formMain.Lvflag = "修正";
            this.Close();

        }

        private void 削除btn_Click(object sender, EventArgs e)
        {
            ValuesAttach();

            //住所録呼出
            //dynamic xlApp = new Excel.Application();
            dynamic xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            dynamic xlBooks = xlApp.Workbooks;
            dynamic xlBook = xlBooks.Open(水利関係住所録WB);
            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(水利関係住所録WB);


            // シートを選択
            var 口酒井名簿 = xlBook.Sheets["口酒井名簿"];

            //// セルの領域を選択
            var TableRange = 口酒井名簿.Range["A2"].CurrentRegion;

            // 選択した領域の値をメモリー上に格納
            object[,] values = TableRange.Value;


            for (int i = 2; i <= values.GetLength(0); i++)
            {
                double セル値 = (double)values[i, 1];

                if (セル値 == double.Parse(Values[0]))
                {
                    DialogResult result = MessageBox.Show("ID　"+ (String)Values[0] + "\n"
                        + "氏名　" + (String)Values[1] + "\n"   //氏名
                        + "郵便番号　" + (String)Values[2] + "\n"   //郵便番号
                        + "住所　" + (String)Values[3] + "\n"
                        + "分類　" + (String)Values[4] + "\n"
                        + "さんのアドレスを削除して良いですか？"
                        , "アドレス削除の確認"
                        , MessageBoxButtons.OKCancel
                        , MessageBoxIcon.Question
                        , MessageBoxDefaultButton.Button2);

                    if (result == DialogResult.Cancel)
                    {
                        MessageBox.Show("削除を中止します。");
                        //Excelのクローズ
                        xlBook.Close();
                        xlBooks.Close();
                        xlApp.Quit();

                        this.Close();

                        return;
                    }

                    口酒井名簿.range[i +":" + i].Delete(-4162);

                    xlBook.Save();

                    IDBox.Text = "";
                    氏名Box.Text = "";
                    郵便番号Box.Text = "";
                    住所Box.Text = "";
                    分類combo.Text = "";

                    ValuesAttach();
                    MessageBox.Show("削除しました");
                    break;
                }
            }

            // 使用したCOMオブジェクトを解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(口酒井名簿);

            //Excelのクローズ
            xlBook.Close();
            xlBooks.Close();
            xlApp.Quit();

            //// 使用したCOMオブジェクトを解放その２
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

            formMain.Lvflag = "削除";

            this.Close();

        }

        private void 住所氏名編集Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Excelのクローズ
            //xlBook.Close();
            //xlBooks.Close();
            //xlApp.Quit();

            //// 使用したCOMオブジェクトを解放その２
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
        }


    }
}
