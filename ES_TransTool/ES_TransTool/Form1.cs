using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Data.OleDb;

namespace ES_TransTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 保存从excel中读取的翻译源
        /// </summary>
        public Dictionary<string, string> m_dicSouce = new Dictionary<string, string>();
        public DataTable dt = new DataTable();

        /// <summary>
        /// 根据字典内容设置ts文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            //把datatable中的数据保存到m_dicSouce中
            int nCount = dt.Rows.Count;
            for (int i = 0; i < nCount; ++i)
            {
                string strKey = dt.Rows[i][comboBox2.SelectedIndex].ToString();
                string strValue = dt.Rows[i][comboBox1.SelectedIndex].ToString();
                m_dicSouce[strKey] = strValue;
            }

            int nSucessNum = 0;
            //string xmlpath = System.AppDomain.CurrentDomain.BaseDirectory + "LocalMaintSystem_en.ts";
            string xmlpath = string.Empty;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "TS文件(*.xml;*.ts)|*.xml;*.ts|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            ofd.Title = "请选择您要更新的TS文件";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                xmlpath = ofd.FileName;
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(xmlpath);
            XmlNode xnTs = doc.SelectSingleNode("//TS");
            XmlNodeList xnTsList = xnTs.ChildNodes;
            foreach (XmlNode xnTsChild in xnTsList)
            {
                if (xnTsChild.Name == "context")
                {
                    XmlNodeList xnList = xnTsChild.ChildNodes;
                    foreach (XmlNode xnf in xnList)
                    {
                        XmlElement xe = (XmlElement)xnf;
                        if (xe.Name == "message")
                        {
                            XmlNodeList xnf1 = xe.ChildNodes;
                            foreach (XmlNode xn2 in xnf1)
                            {
                                switch (comboBox2.SelectedIndex)
                                {
                                    case 0:
                                        {
                                            //用中文替换，使用source节点
                                            if (xn2.Name.CompareTo("source") == 0)
                                            {
                                                string sourceText = xn2.InnerText;
                                                XmlNode xnTranslation = xn2.NextSibling;
                                                if (xnTranslation.Name == "translation")
                                                {
                                                    if (m_dicSouce.ContainsKey(sourceText))
                                                    {
                                                        //修改节点translation的值
                                                        xnTranslation.InnerText = m_dicSouce[sourceText];
                                                        nSucessNum++;
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                    case 1:
                                        {
                                            //用英文替换，使用translation节点
                                            if (xn2.Name.CompareTo("translation") == 0)
                                            {
                                                string translationText = xn2.InnerText;
                                                if (m_dicSouce.ContainsKey(translationText))
                                                {
                                                    //修改节点translation的值
                                                    xn2.InnerText = m_dicSouce[translationText];
                                                    nSucessNum++;
                                                }
                                            }
                                        }
                                        break;
                                }
                            }
                        }

                    }
                }
            }
            doc.Save(xmlpath);
            MessageBox.Show(nSucessNum.ToString() + "elements trans success!");
        }


        /// <summary>
        /// 从excel中获取需要翻译的资源，保存到DataTable
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static DataTable ReadExcelToTable(string path)
        {
            try
            {
                string strConn;
                // strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=YES\"";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet1$]";//可是更改Sheet名称，比如sheet2，等等

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, "Sheet1");
                OleConn.Close();

                return OleDsExcle.Tables["Sheet1"];
            }
            catch (Exception err)
            {
                MessageBox.Show("数据绑定Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }


        /// <summary>
        /// 选择翻译文件excel并加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            ofd.Title = "请选择翻译源文件（EXCEL）";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                //加载excel
                // DataTable dt;
                dt = ReadExcelToTable(ofd.FileName);
                dataGridView1.DataSource = dt;

                //把第一行语言项加到列表
                int nLanguageCound = dt.Columns.Count;
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                for (int i = 0; i < nLanguageCound; ++i)
                {
                    comboBox1.Items.Add(dt.Columns[i].ToString());
                    comboBox2.Items.Add(dt.Columns[i].ToString());
                }
                comboBox1.SelectedIndex = 1;
                comboBox2.SelectedIndex = 0;
            }
        }
    }
}
