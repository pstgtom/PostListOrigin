﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;

namespace 口酒井農業水利組合郵送会員住所録
{
    public partial class 住所氏名編集Form : Form
    {
        private string[] Values;
        public 住所録リストForm formMain;
        private string _mode;
        public string サーバ;

        public 住所氏名編集Form()
        {
            InitializeComponent();
            分類セット();
            ValuesAttach();
            formMain.接続文字列設定();
        }

        public void 処理モード(string 押下げボタン)
        {
            if (押下げボタン == "追加")
            {
                新規追加btn.Enabled = true;
            }
            else
            {
                新規追加btn.Enabled = false;
            }

        }

        private void 分類セット()
        {
            分類combo.Items.Add("入作");
            分類combo.Items.Add("企業（振込）");
            分類combo.Items.Add("墓地管理");

            分類combo.Text = "入作";
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
        }


        private void クリアbtn_Click(object sender, EventArgs e)
        {
            this.IDBox.Text = "";
            this.氏名Box.Text = "";
            this.郵便番号Box.Text = "";
            this.住所Box.Text = "";
            this.分類combo.Text = "";

            this.新規追加btn.Enabled = true;

        }



        private void 新規追加btn_Click(object sender, EventArgs e)
        {
            if (分類combo.SelectedIndex == -1)
            {
                MessageBox.Show("分類が選択されていません。");
                return;
            }


            ValuesAttach();

            NpgsqlConnection myCon = new NpgsqlConnection("Server=" + サーバ + ";Port=5432;Uid=kuchisakai;Pwd=9mei5jikai#;Database=test9meidb;");
            myCon.Open();

            string SQLstr = "SELECT 氏名, 住所, 郵便番号, 分類 FROM 郵送名簿 WHERE 氏名 = '" + Values[1] + "' AND 分類 = '" + Values[4] + "'";

            NpgsqlCommand command = new NpgsqlCommand(SQLstr, myCon);
            NpgsqlDataReader dr = command.ExecuteReader();

            if (dr.HasRows == true)
            {
                MessageBox.Show(Values[1] + "(" + Values[4] + ") さんは既に登録されています。\n\n"
                                    + "新規に追加する場合は今とは別の区分としてください。");
                myCon.Close();
                return;
            }

            var ans = MessageBox.Show("新規に郵送メンバーを追加して良いですか？"
                                        ,"郵送メンバーの新規追加"
                                        , MessageBoxButtons.YesNo
                                        ,MessageBoxIcon.Question);
            switch (ans)
            {
                case DialogResult.No:
                    MessageBox.Show("中止します。");
                    break;
            }


            myCon.Close();
            myCon.Open();

            SQLstr = "INSERT INTO owner(所有者,住所,郵便番号,区分id) " +
                                    "VALUES ('" + Values[1] + "', '" + Values[3] + "', '" + Values[2] + "', " + 
                                           " (SELECT id FROM trait WHERE 区分 = '" + Values[4] + "'))";

            command = new NpgsqlCommand(SQLstr, myCon);
            command.ExecuteNonQuery();


            MessageBox.Show("追加しました");

            this.Close();
        }


        private void 修正btn_Click(object sender, EventArgs e)
        {
            var ans = MessageBox.Show("修正して良いですか？"
                            , "修正"
                            , MessageBoxButtons.YesNo
                            , MessageBoxIcon.Question);
            switch (ans)
            {
                case DialogResult.No:
                    MessageBox.Show("中止します。");
                    break;
            }


            ValuesAttach();

            NpgsqlConnection myCon = new NpgsqlConnection("Server=" + サーバ + ";Port=5432;Uid=kuchisakai;Pwd=9mei5jikai#;Database=test9meidb;");
            myCon.Open();
            string SQLstr = "UPDATE owner SET 所有者 ='" + Values[1] + "', " +
                                             "住所 = '" + Values[3] + "', " + 
                                             "郵便番号 = '" + Values[2] + "', " +
                                             "区分id = (SELECT id FROM trait WHERE 区分 = '" + Values[4] + "')" +
                                        "WHERE id = " + int.Parse(Values[0]); 

            NpgsqlCommand command = new NpgsqlCommand(SQLstr, myCon);
            command.ExecuteNonQuery();

            MessageBox.Show("修正しました");
 
            this.Close();

        }

        private void 削除btn_Click(object sender, EventArgs e)
        {
            //ValuesAttach();

            //myCon.Open();
            //var transa = myCon.BeginTransaction();

            //string SQLstr = "DELETE FROM owner WHERE id = " + int.Parse(Values[0]);
            //NpgsqlCommand command = new NpgsqlCommand(SQLstr, myCon);
            //command.ExecuteNonQuery();

            //var ans = MessageBox.Show("削除して良いですか？"
            //    , "修正"
            //    , MessageBoxButtons.YesNo
            //    , MessageBoxIcon.Question);
            //switch (ans)
            //{
            //    case DialogResult.No:
            //        MessageBox.Show("中止します。");
            //        transa.Rollback();
            //        break;
            //}

            //transa.Commit();
            //MessageBox.Show("削除しました");

            //this.Close();

        }

        private void 住所氏名編集Form_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

    }
}
