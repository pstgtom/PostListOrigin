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
using Npgsql;

namespace 口酒井農業水利組合郵送会員住所録
{
    public partial class 住所録リストForm : Form
    {

        //修正後のデータ受け取り配列変数
        住所氏名編集Form fs;
        NpgsqlConnection myCon = new NpgsqlConnection("Server=fertila;Port=5432;Uid=kuchsakai;Pwd=9mei5jikai#;Database=test9meidb");

        public 住所録リストForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            Text = "郵送住所録フォーム　" + "Ver.1.02" + "　　　　　　　定型長３封筒に印刷してください。";

            ListView1呼出();

            //NpgsqlConnection myCon = new NpgsqlConnection("Server=fertila;Port=5432;Uid=kuchsakai;Pwd=9mei5jikai#;Database=test9meidb");
            myCon.Open();

            NpgsqlCommand SQLstr = new NpgsqlCommand("SELECT * FROM sender", myCon);

            NpgsqlDataReader dr = SQLstr.ExecuteReader();

                int i;

                try
                {

                    while (dr.Read())
                    {
                        for (i = 0; i < dr.FieldCount; i++)
                        {
                            Console.Write("{0} \t", dr[i]);
                        }
                        Console.WriteLine();
                    }
                }

                finally
                {
                    myCon.Close();
                }

                //差出人Box.Text = dataReader(1);
                //差出人住所Box.Text = dataReader[2];

            MessageBox.Show("定型長３封筒に印刷してください。");

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

            // 配列アクセスができるので、それぞれをリストビューに追加

            ListViewItem lvi;

            myCon.Open();

            try
            {
                NpgsqlCommand SQLstr = new NpgsqlCommand("SELECT * FROM 郵送名簿", myCon);

                NpgsqlDataReader dr = SQLstr.ExecuteReader();

                int i;

                while (dr.Read())
                {
                    for (i = 0; i < dr.FieldCount; i++)
                        {
                            Console.Write("{0} \t", dr[i]);
                        }
                        Console.WriteLine();
                //    lvi = listView1.Items.Add(dr[i].ToString());
                //    lvi.SubItems.Add(dr[1].ToString());
                //    lvi.SubItems.Add(dr[2].ToString());
                //    lvi.SubItems.Add(dr[3].ToString());
                //    lvi.SubItems.Add(dr[4].ToString());
                }
            }
            finally
            {
                myCon.Close();
            }

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
            //string 氏名, 郵便番号, 住所;

            //object[,] values = TableRange.Value;


            //for (int i = 2; i <= values.GetLength(0); i++)
            //{
            //    string セル値 = (String)values[i, 5];

            //    if (セル値 == 分類 || 分類 == "全件")
            //    {
            //        氏名 = (String)values[i, 2];
            //        郵便番号 = (String)values[i, 3];
            //        住所 = (String)values[i, 4];

            //        宛名印刷(氏名, 郵便番号, 住所);
            //    }
            //}

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

            //宛名面.Shapes(1).Name = "宛先郵便番号";
            //宛名面.Shapes(2).Name = "宛先住所";
            //宛名面.Shapes(3).Name = "宛先氏名1";
            //宛名面.Shapes(4).Name = "宛先氏名2";
            //宛名面.Shapes(5).Name = "差出人氏名";
            //宛名面.Shapes(6).Name = "差出人住所";


            //宛名面.Shapes("宛先郵便番号").TextFrame.Characters.Text = "";      //宛先郵便番号
            //宛名面.Shapes("宛先住所").TextFrame.Characters.Text = "";          //宛先住所
            //宛名面.Shapes("宛先氏名1").TextFrame.Characters.Text = "";         //宛先氏名1
            //宛名面.Shapes("宛先氏名2").TextFrame.Characters.Text = "";         //宛先氏名2
            //宛名面.Shapes("差出人氏名").TextFrame.Characters.Text = "";       //差出人氏名
            //宛名面.Shapes("差出人住所").TextFrame.Characters.Text = "";       //差出人住所


        }


        private void 宛名印刷(string 氏名, string 郵便番号, string 住所)
        {
            初期化();

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

            住所.Replace("－", "ー");
            住所.Replace("-", "ー");

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


            //宛名面.Shapes("宛先郵便番号").TextFrame.Characters.Text = 郵便番号;                           //宛先郵便番号
            //宛名面.Shapes("宛先住所").TextFrame.Characters.Text = 住所;                  //宛先住所
            //宛名面.Shapes("宛先氏名1").TextFrame.Characters.Text = 氏名;                             //宛先氏名1
            //宛名面.Shapes("宛先氏名2").TextFrame.Characters.Text = 氏名2;                                           //宛先氏名2
            //宛名面.Shapes("差出人氏名").TextFrame.Characters.Text = 差出人氏名;                         //差出人氏名
            //宛名面.Shapes("差出人住所").TextFrame.Characters.Text = 差出人住所;       //差出人住所

            //dynamic printOut = 宛名面.Range["A1:D32"].PrintOut;

        }



        private void 差出人変更btn_Click(object sender, EventArgs e)
        {


            MessageBox.Show("変更を保存しました");


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
        }

 
        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            一軒編集btn_Click(sender, e);
        }
    }
}



