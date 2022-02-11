using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.VisualBasic.FileIO;

namespace OpenBirdsOperation

{
    public partial class commandList : Form
    {
        public operationMain formMain;
        public commandList()
        {
            InitializeComponent();
        }

        public void addCommand(string command, string comments, string remark) {
            missingPacketDataGridView.Rows.Add(command, comments, remark);
            tabControl1.SelectedTab = missingTabPage;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "csv file(*.csv)|*.csv|all file(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Title = "Select command file";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(ofd.FileName);

                TextFieldParser parser = new TextFieldParser(ofd.FileName, Encoding.GetEncoding("Shift_JIS"));
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");


                //データをすべてクリア
                if (tabControl1.SelectedIndex == 0)
                {
                    dataGridView2.Rows.Clear();
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    dataGridView2.Rows.Clear();
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    dataGridView3.Rows.Clear();
                }
                else
                {
                    missingPacketDataGridView.Rows.Clear();
                }

                parser.ReadFields();    //1行目読み飛ばし
                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields(); // 1行読み込み
                                                        // 読み込んだデータ(1行をDataGridViewに表示する)

                    if (tabControl1.SelectedIndex == 0) {
                        dataGridView1.Rows.Add(row);
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        dataGridView2.Rows.Add(row);
                    }
                    else if (tabControl1.SelectedIndex == 2)
                    {
                        dataGridView3.Rows.Add(row);
                    }
                    else
                    {
                        missingPacketDataGridView.Rows.Add(row);
                    }

                }

            }

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (formMain != null)
            {
                if (tabControl1.SelectedIndex == 0 && dataGridView1.CurrentRow.Cells[0].Value == null) {
                    return;
                }
                else if (tabControl1.SelectedIndex == 1 && dataGridView2.CurrentRow.Cells[0].Value == null)
                {
                    return;
                }
                else if (tabControl1.SelectedIndex == 2 && dataGridView3.CurrentRow.Cells[0].Value == null)
                {
                    return;
                }
                else if (tabControl1.SelectedIndex == 3 && missingPacketDataGridView.CurrentRow.Cells[0].Value == null)
                {
                    return;
                }

                if (tabControl1.SelectedIndex == 0)
                {
                    formMain.CommandString = dataGridView1.CurrentRow.Cells[0].Value.ToString();
                    formMain.CommentsString = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex + 1].Cells[0];
                }
                else if (tabControl1.SelectedIndex == 1)
                {
                    formMain.CommandString = dataGridView2.CurrentRow.Cells[0].Value.ToString();
                    formMain.CommentsString = dataGridView2.CurrentRow.Cells[1].Value.ToString();
                    dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex + 1].Cells[0];
                }
                else if (tabControl1.SelectedIndex == 2)
                {
                    formMain.CommandString = dataGridView3.CurrentRow.Cells[0].Value.ToString();
                    formMain.CommentsString = dataGridView3.CurrentRow.Cells[1].Value.ToString();
                    dataGridView3.CurrentCell = dataGridView3.Rows[dataGridView3.CurrentCell.RowIndex + 1].Cells[0];
                }
                else
                {
                    formMain.CommandString = missingPacketDataGridView.CurrentRow.Cells[0].Value.ToString();
                    formMain.CommentsString = missingPacketDataGridView.CurrentRow.Cells[1].Value.ToString();
                    missingPacketDataGridView.CurrentCell = missingPacketDataGridView.Rows[missingPacketDataGridView.CurrentCell.RowIndex + 1].Cells[0];
                }

            }
        }

        private void DataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Button3_Click(sender, e);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "csv file(*.csv)|*.csv|all file(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Title = "Select command file";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(ofd.FileName);

                TextFieldParser parser = new TextFieldParser(ofd.FileName, Encoding.GetEncoding("Shift_JIS"));
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");


                parser.ReadFields();    //1行目読み飛ばし
                while (!parser.EndOfData)
                {
                    string[] row = parser.ReadFields(); // 1行読み込み
                                                        // 読み込んだデータ(1行をDataGridViewに表示する)

                    if (tabControl1.SelectedIndex == 0)
                    {
                        dataGridView1.Rows.Add(row);
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        dataGridView2.Rows.Add(row);
                    }
                    else if (tabControl1.SelectedIndex == 2)
                    {
                        dataGridView3.Rows.Add(row);
                    }
                    else
                    {
                        missingPacketDataGridView.Rows.Add(row);
                    }

                }

            }

        }

        private void CommandList_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 3;
        }
    }
}
