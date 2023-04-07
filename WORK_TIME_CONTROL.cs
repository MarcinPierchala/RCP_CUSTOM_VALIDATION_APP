using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;

namespace WORK_TIME_CTRL
{
    public partial class formMainWindow : Form
    {
        
        public formMainWindow()
        {
            InitializeComponent();
        }
        private void btopen_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();  //create openfileDialog Object
                openFileDialog1.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
                openFileDialog1.FilterIndex = 3;

                openFileDialog1.Multiselect = false;        //not allow multiline selection at the file selection level
                openFileDialog1.Title = "Open Text File-R13";   //define the name of openfileDialog
                openFileDialog1.InitialDirectory = @"Desktop"; //define the initial directory

                if (openFileDialog1.ShowDialog() == DialogResult.OK)        //executing when file open
                {
                    string pathName = openFileDialog1.FileName;
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    DataTable tbContainer = new DataTable();
                    string strConn = string.Empty;
                    string sheetName = "Sheet1";

                    FileInfo file = new FileInfo(pathName);
                    if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
                    string extension = file.Extension;
                    switch (extension)
                    {
                        case ".xls":
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                        case ".xlsx":
                            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                            break;
                        default:
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                    }
                    OleDbConnection cnnxls = new OleDbConnection(strConn);
                    OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
                    oda.Fill(tbContainer);

                    dataGridView1.DataSource = tbContainer;
                    //dataGridView1.AutoResizeColumns();
                    dataGridView1.EndEdit();
                    dataGridView1.Columns.Add("ALERT1", "ALERT2");

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        try
                        {
                            if ((Convert.ToInt32(row.Cells[0].Value) > 0 && (Convert.ToInt32(row.Cells[0].Value) < 31)))
                            {
                                if (row.Cells[1].Value.ToString() == "Niedziela i Święta" || row.Cells[1].Value.ToString() == "Wolne")
                                {
                                    try
                                    {
                                        row.Cells[12].Value = Convert.ToDateTime(row.Cells[3].Value) - Convert.ToDateTime(row.Cells[2].Value);// "UWAGA SWIETO";
                                        row.Cells[12].Style.BackColor = Color.Red;
                                    }
                                    catch { }

                                }
                                else if (row.Cells[1].Value.ToString() == "Pracujący")
                                {
                                    try
                                    {
                                        if ((Convert.ToDateTime(row.Cells[2].Value) > Convert.ToDateTime("05:30") && Convert.ToDateTime(row.Cells[3].Value) < Convert.ToDateTime("14:30"))
                                            || (Convert.ToDateTime(row.Cells[2].Value) > Convert.ToDateTime("13:30") && Convert.ToDateTime(row.Cells[3].Value) < Convert.ToDateTime("22:30")))
                                        {
                                            //row.Cells[12].Value = Convert.ToDateTime("08:00");
                                            row.Cells[12].Value = Convert.ToDateTime(row.Cells[3].Value) - Convert.ToDateTime(row.Cells[2].Value);
                                            row.Cells[12].Style.BackColor = Color.Green;
                                        }
                                        else if (row.Cells[2].Value == null || row.Cells[2].Value == DBNull.Value || String.IsNullOrWhiteSpace(row.Cells[2].Value.ToString()))
                                        {
                                            row.Cells[12].Style.BackColor = Color.HotPink;
                                            row.Cells[12].Value = "!!!UWAGA!!!";
                                        }
                                        else
                                        {
                                            row.Cells[12].Value = row.Cells[12].Value = Convert.ToDateTime(row.Cells[3].Value) - Convert.ToDateTime(row.Cells[2].Value);//"!!NADGODZINY!!";
                                            row.Cells[12].Style.BackColor = Color.Lime;
                                        }
                                    }
                                    catch
                                    {

                                    }
                                }
                                if(Convert.ToDateTime(row.Cells[4].Value) < Convert.ToDateTime("08:00"))
                                {
                                    row.Cells[4].Style.BackColor = Color.Red;
                                    row.Cells[4].Style.ForeColor = Color.Yellow;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error!");
            }
        }

        private void AboutMe_Click(object sender, EventArgs e)
        {
            MessageBox.Show("To jest na (prawie) szybko naszkryfany programik\ndo sprawdzania poprawności wypełnienia druku RCP\nAutor: Marcin Pierchała\nmarcin.pierchala@icloud.com");
        }

        private void Exit_click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Dzięki za skorzystanie z programiku :-)\nChcesz zamknąć program?",
                      "Chcesz zamknąć program?", MessageBoxButtons.YesNo);
            switch (dr)
            {
                case DialogResult.Yes:
                    this.Close();
                    break;
                case DialogResult.No:
                    break;
            }
        }       
    }
}
