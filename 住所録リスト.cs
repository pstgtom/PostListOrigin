//using Microsoft.Office.Core;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

namespace 口酒井農業水利組合郵送会員住所録
{
    public partial class 住所録リストForm : Form
    {
        //住所録ファイルのパス
        //public string path = (@"D:\develope\口酒井\住所録\");
        public string 水利関係住所録WB = (@"C:\dropbox\住所録\水利関係住所録.xlsx");
        //public string 水利関係住所録WB = ("水利関係住所録.xlsx");

        //住所録呼出
        public dynamic xlApp;
        public dynamic xlBooks;
        public dynamic xlBook;

        //修正後のデータ受け取り配列変数
        住所氏名編集Form fs;


        public 住所録リストForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //水利関係住所録WBが開かれてないかのチェック
            try
            {
                var FA = File.AppendText(水利関係住所録WB);
                FA.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(水利関係住所録WB + "は既に開いています。閉じてからやり直してください。");
                Close();
                return;
            }

            Text = "郵送住所録フォーム　" + "Ver.1.02" + "　　　　　　　定型長３封筒に印刷してください。";

            //住所録呼出
            //xlApp = new Excel.Application();
            xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            xlBooks = xlApp.Workbooks;
            xlBook = xlBooks.Open(水利関係住所録WB);

            ListView1呼出();

            // シートを選択
            var sashidasinin = xlBook.Sheets["差出人"];

            // セルの領域を選択
            var 差出人 = sashidasinin.Range["A2", "B2"];

            // 選択した領域の値をメモリー上に格納
            object[,] Sender = 差出人.Value;

            // 選択した領域の値をメモリー上に格納
            差出人Box.Text = (string)Sender[1, 1];
            差出人住所Box.Text = (string)Sender[1, 2];


            // 使用したCOMオブジェクトを解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(差出人);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sashidasinin);

            // Excelのクローズ
            //xlBook.Saved = true;
            //xlBook.Close();
            //xlBooks.Close();
            //xlApp.Quit();

            MessageBox.Show("定型長３封筒に印刷してください。");

            //// 使用したCOMオブジェクトを解放その２
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

        }


        public void ListView1呼出()
        {
            listView1.Clear();
            // ListViewコントロールのプロパティを設定
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            //listView1.Sorting = SortOrder.Ascending;
            listView1.View = View.Details;
            //リストのカラム登録
            listView1.Columns.Add("id", 50);
            listView1.Columns.Add("氏　名", 200);
            listView1.Columns.Add("〒", 60);
            listView1.Columns.Add("住　所", 300);
            listView1.Columns.Add("分類", 75);

            // シートを選択
            var sheet = xlBook.Sheets["口酒井名簿"];

            // セルの領域を選択
            //var TableRange = sheet.Range["A2", "E47"];
            var TableRange = sheet.Range["A2"].CurrentRegion;

            // 選択した領域の値をメモリー上に格納
            object[,] values = TableRange.Value;

            // 配列アクセスができるので、それぞれをリストビューに追加

            ListViewItem lvi;

            for (int i = 2; i <= values.GetLength(0); i++)
            {
                lvi = listView1.Items.Add((String)values[i, 1].ToString());
                lvi.SubItems.Add((string)values[i, 2]);
                lvi.SubItems.Add((string)values[i, 3]);
                lvi.SubItems.Add((string)values[i, 4]);
                lvi.SubItems.Add((string)values[i, 5]);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);

        }




        private void 全件印刷btn_Click(object sender, EventArgs e)
        {
            印刷対象選択("全件");
        }

        private void 入作関係btn_Click(object sender, EventArgs e)
        {
            印刷対象選択("入作関係");

        }

        private void 企業協力金btn_Click(object sender, EventArgs e)
        {
            印刷対象選択("企業協力金");

        }

        private void 墓地管理btn_Click(object sender, EventArgs e)
        {
            印刷対象選択("墓地管理");

        }


        private void 印刷対象選択(string 分類)
        {
            string 氏名, 郵便番号, 住所;

            //xlBook = xlBooks.Open(水利関係住所録WB);

            // シートを選択
            var 口酒井名簿 = xlBook.Sheets["口酒井名簿"];

            // セルの領域を選択
            var TableRange = 口酒井名簿.Range["A2"].CurrentRegion;

            // 選択した領域の値をメモリー上に格納
            object[,] values = TableRange.Value;


            for (int i = 2; i <= values.GetLength(0); i++)
            {
                string セル値 = (String)values[i, 5];

                if (セル値 == 分類 || 分類 == "全件")
                {
                    氏名 = (String)values[i, 2];
                    郵便番号 = (String)values[i, 3];
                    住所 = (String)values[i, 4];

                    宛名印刷(氏名, 郵便番号, 住所);
                }
            }

            // 使用したCOMオブジェクトを解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(口酒井名簿);

        }


        //一軒ずつ印刷
        private void 一軒印刷btn_Click(object sender, EventArgs e)
        {
            // 選択項目があるかどうかを確認する
            if (listView1.SelectedItems.Count == 0)
            {
                // 選択項目がないので処理をせず抜ける
                MessageBox.Show("宛名印刷する方が選択されていません。");
                return;
            }

            // 選択項目を取得する
            ListViewItem itemx = listView1.SelectedItems[0];
            var ans = MessageBox.Show(itemx.SubItems[4].Text + Environment.NewLine + Environment.NewLine
                                    + itemx.SubItems[3].Text + Environment.NewLine + Environment.NewLine
                                    + itemx.SubItems[1].Text + " さん " + Environment.NewLine + Environment.NewLine
                                    + "の宛名印刷をしますか？"
                                , "一軒ごと宛名印刷の確認"
                                , MessageBoxButtons.YesNo
                                , MessageBoxIcon.Question);

            switch (ans)
            {
                case DialogResult.No:
                    MessageBox.Show("一軒印刷を中止します。");
                    break;
                case DialogResult.Yes:
                    初期化();
                    宛名印刷(itemx.SubItems[1].Text, itemx.SubItems[2].Text, itemx.SubItems[3].Text);
                    break;
            }
        }


        private void 初期化()
        {
            //xlBook = xlBooks.Open(水利関係住所録WB);

            ////シートを選択
            var 宛名面 = xlBook.Sheets["宛名面"];

            宛名面.Shapes(1).Name = "宛先郵便番号";
            宛名面.Shapes(2).Name = "宛先住所";
            宛名面.Shapes(3).Name = "宛先氏名1";
            宛名面.Shapes(4).Name = "宛先氏名2";
            宛名面.Shapes(5).Name = "差出人氏名";
            宛名面.Shapes(6).Name = "差出人住所";


            宛名面.Shapes("宛先郵便番号").TextFrame.Characters.Text = "";      //宛先郵便番号
            宛名面.Shapes("宛先住所").TextFrame.Characters.Text = "";          //宛先住所
            宛名面.Shapes("宛先氏名1").TextFrame.Characters.Text = "";         //宛先氏名1
            宛名面.Shapes("宛先氏名2").TextFrame.Characters.Text = "";         //宛先氏名2
            宛名面.Shapes("差出人氏名").TextFrame.Characters.Text = "";       //差出人氏名
            宛名面.Shapes("差出人住所").TextFrame.Characters.Text = "";       //差出人住所


        }


        private void 宛名印刷(string 氏名, string 郵便番号, string 住所)
        {
            初期化();

            //xlBook = xlBooks.Open(水利関係住所録WB);

            ////シートを選択
            var 宛名面 = xlBook.Sheets["宛名面"];

            //郵便番号がnullの場合、0000でnull合体演算をする
            string fourZeros = "000-0000";
            string 住所なし = "　";
            string 名無し = "　";

            郵便番号 = 郵便番号 ?? fourZeros;
            住所 = 住所 ?? 住所なし;
            氏名 = 氏名 ?? 名無し;
            string 氏名2 = "";
            string 差出人氏名 = this.差出人Box.Text;
            string 差出人住所 = this.差出人住所Box.Text;


            郵便番号 = 郵便番号.Replace("-", "");
            郵便番号 = 郵便番号.Replace("－", "");
            郵便番号 = 郵便番号.Replace("ー", "");

            //MessageBox.Show(郵便番号);
            郵便番号 = 郵便番号.Replace('0', '０');
            郵便番号 = 郵便番号.Replace('1', '１');
            郵便番号 = 郵便番号.Replace('2', '２');
            郵便番号 = 郵便番号.Replace('3', '３');
            郵便番号 = 郵便番号.Replace('4', '４');
            郵便番号 = 郵便番号.Replace('5', '５');
            郵便番号 = 郵便番号.Replace('6', '６');
            郵便番号 = 郵便番号.Replace('7', '７');
            郵便番号 = 郵便番号.Replace('8', '８');
            郵便番号 = 郵便番号.Replace('9', '９');
            //MessageBox.Show(郵便番号);

            //宛名面2.Range["L1"].value = 郵便番号.Substring(0, 1);
            //宛名面2.Range["E1"].value = 郵便番号.Substring(1, 1);
            //宛名面2.Range["F1"].value = 郵便番号.Substring(2, 1);
            //宛名面2.Range["G1"].value = 郵便番号.Substring(3, 1);
            //宛名面2.Range["I1"].value = 郵便番号.Substring(4, 1);
            //宛名面2.Range["J1"].value = 郵便番号.Substring(5, 1);
            //宛名面2.Range["K1"].value = 郵便番号.Substring(6, 1);

            住所.Replace("－", "ー");
            住所.Replace("-", "ー");

            //宛名面2.Range["J3"].value = 住所;

            氏名 = 氏名.Replace("　", "");
            氏名 = 氏名.Replace(" ", "");

            int len = 氏名.Length;

            if (len >= 8)
            {
                if (氏名.Contains("関西電力送配電㈱"))
                {
                    //宛名面2.Range["D3"].value = "関西電力送配電㈱";
                    氏名 = "関西電力送配電㈱";
                    //宛名面2.Range["C3"].value = "兵庫支社神戸電力本部尼崎電力所御中";
                    氏名2 = "兵庫支社神戸電力本部尼崎電力所御中";
                }
            }
            else if (len == 3)
            {
                氏名 = 氏名.Insert(2, "　") + "様";
            }
            else
            {
                氏名 = 氏名 + "様";
            }

            //宛名面2.Range["D3"].value = 氏名;


            宛名面.Shapes("宛先郵便番号").TextFrame.Characters.Text = 郵便番号;                           //宛先郵便番号
            宛名面.Shapes("宛先住所").TextFrame.Characters.Text = 住所;                  //宛先住所
            宛名面.Shapes("宛先氏名1").TextFrame.Characters.Text = 氏名;                             //宛先氏名1
            宛名面.Shapes("宛先氏名2").TextFrame.Characters.Text = 氏名2;                                           //宛先氏名2
            宛名面.Shapes("差出人氏名").TextFrame.Characters.Text = 差出人氏名;                         //差出人氏名
            宛名面.Shapes("差出人住所").TextFrame.Characters.Text = 差出人住所;       //差出人住所

            //dynamic printOut = 宛名面2.Range["A1:L3"].PrintOut;
            dynamic printOut = 宛名面.Range["A1:D32"].PrintOut;

            //// Excelのクローズ
            //xlBook.Saved = true;
            //xlBook.Close();
            //xlApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
        }



        private void 差出人変更btn_Click(object sender, EventArgs e)
        {

            //xlApp = new Excel.Application();
            //xlBooks = xlApp.Workbooks;
            //xlBook = xlBooks.Open(水利関係住所録WB);

            //MessageBox.Show(xlApp.name);
            ////MessageBox.Show(xlBooks.thisworkbook.name);
            //MessageBox.Show(xlBook.name);

            ////シートを選択
            var Sashidashinin = xlBook.Sheets["差出人"];

            //if (Sashidashinin.Range["A2"].value = 差出人Box.Text)

            Sashidashinin.Range["A2"].value = 差出人Box.Text;
            Sashidashinin.Range["B2"].value = 差出人住所Box.Text;

            //上書き保存
            xlBook.Save();

            MessageBox.Show("変更を保存しました");

            // Excelのクローズ
            //xlBook.Close();
            //xlApp.Quit();

            ////使用したCOMオブジェクトを解放その２
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);

        }


        private void 一軒編集btn_Click(object sender, EventArgs e)
        {
            // 選択項目があるかどうかを確認する
            if (listView1.SelectedItems.Count == 0)
                {
                    // 選択項目がないので処理をせず抜ける
                    MessageBox.Show("住所氏名を編集する方が選択されていません。");
                    return;
                }

                fs = new 住所氏名編集Form();
                fs.formMain = this;

            // Excelのクローズ
                xlBook.Saved = true;
                xlBook.Close();
                //xlApp.Quit();

                ListViewItem itemx = listView1.SelectedItems[0];

                string ID = itemx.SubItems[0].Text;
                string 住所 = itemx.SubItems[1].Text;
                string 郵便番号 = itemx.SubItems[2].Text;
                string 氏名 = itemx.SubItems[3].Text;
                string 分類 = itemx.SubItems[4].Text;

                string[] 修正前データ = { ID, 住所, 郵便番号, 氏名, 分類};

                fs.AddSet = 修正前データ;

                //住所氏名変更Formを呼び出す
                fs.ShowDialog();


            //xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            //xlBooks = xlApp.Workbooks;
            xlBook = xlBooks.Open(水利関係住所録WB);


            if (Lvflag == "追加")
            {
                //ListViewに一軒追加();
                ListViewItem itemLast = new ListViewItem();

                string[] 修正後データ = fs.AddSet;
                itemLast.Text = 修正後データ[0];
                itemLast.SubItems.Add(修正後データ[1]);
                itemLast.SubItems.Add(修正後データ[2]);
                itemLast.SubItems.Add(修正後データ[3]);
                itemLast.SubItems.Add(修正後データ[4]);

                listView1.Items.Add(itemLast);

            }
            else if (Lvflag == "修正")
            {

                string[] 修正後データ = fs.AddSet;
                itemx.SubItems[0].Text = 修正後データ[0];
                itemx.SubItems[1].Text = 修正後データ[1];
                itemx.SubItems[2].Text = 修正後データ[2];
                itemx.SubItems[3].Text = 修正後データ[3];
                itemx.SubItems[4].Text = 修正後データ[4];

            }
            else if (Lvflag == "削除")
            {
                listView1.Items.Remove(itemx);
                string[] 修正後データ = fs.AddSet;

            }

            Lvflag = "";
            fs.Close();

            }

        private string リストビュー更新;

        public string Lvflag
        {
            set
            {
                リストビュー更新 = value;
            }
            get
            {
                return リストビュー更新;
            }
        }

        private void 住所録リストForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //// 使用したCOMオブジェクトを解放
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(TableRange);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(差出人);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(sashidasinin);

            //Excelのクローズ
            xlBook.Saved = true;
            xlBook.Close();
            xlBooks.Close();
            xlApp.Quit();

            //// 使用したCOMオブジェクトを解放その２
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject((object)xlApp);
        }

 
        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            一軒編集btn_Click(sender, e);
        }
    }
}



