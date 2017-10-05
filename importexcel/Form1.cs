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
using System.Data.SqlClient;
using System.IO;


namespace importexcel
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=DINESH-18;Initial Catalog=POS;Integrated Security=True");
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFilelog1 = new OpenFileDialog();
            if(openFilelog1.ShowDialog()==System.Windows.Forms.DialogResult.OK)
            {
                this.textBox_path.Text = openFilelog1.FileName;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string PathConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBox_path.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(PathConn);

            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from[" + textBox_sheet.Text + "$]", conn);
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int n = dataGridView1.Rows.Add();
            for(int i = 0;i<n;i++)
            {
                String sppID = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                String pID = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                String pNM = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                String HSN = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                String EANcode = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                decimal MFG = Convert.ToDecimal(dataGridView1.Rows[i].Cells[5].Value);
                decimal EXP = Convert.ToDecimal(dataGridView1.Rows[i].Cells[6].Value);
                String UOM = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                decimal QTY = Convert.ToDecimal(dataGridView1.Rows[i].Cells[8].Value);
                decimal sP = Convert.ToDecimal(dataGridView1.Rows[i].Cells[9].Value);
                decimal sPG = Convert.ToDecimal(dataGridView1.Rows[i].Cells[10].Value);
                decimal MRP = Convert.ToDecimal(dataGridView1.Rows[i].Cells[11].Value);
                decimal GST = Convert.ToDecimal(dataGridView1.Rows[i].Cells[12].Value);
                decimal GSTamt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[13].Value);
                decimal CGST = Convert.ToDecimal(dataGridView1.Rows[i].Cells[14].Value);
                decimal CGSTamt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[15].Value);
                decimal SGST = Convert.ToDecimal(dataGridView1.Rows[i].Cells[16].Value);
                decimal SGSTamt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[17].Value);
                decimal UTGST = Convert.ToDecimal(dataGridView1.Rows[i].Cells[18].Value);
                decimal UTGSTamt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[19].Value);
                decimal IGST = Convert.ToDecimal(dataGridView1.Rows[i].Cells[20].Value);
                decimal IGSTamt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[21].Value);
                decimal sPT = Convert.ToDecimal(dataGridView1.Rows[i].Cells[22].Value);
                SqlCommand cmd = new SqlCommand(@"INSERT INTO [POS].[dbo].[Invoice]([suppProductID],[productID],[productName],[HSN],[EANcode],[MFG],[EXP],[UOM],[qty],[supplierPrice],[supplierPriceGST],[MRP],[GST],[GSTamt],[CGST],[CGSTamt],[SGST],[SGSTamt],[UTGST],[UTGSTamt],[IGST],[IGSTamt],[supplier Tax])
     VALUES('" + sppID.ToString() + "','" + pID.ToString() + "','" + pNM.ToString() + "','" + HSN.ToString() + "','" + EANcode.ToString() + "','" + MFG.ToString("YYYY/MM/DD") + "','" + EXP.ToString("YYYY/MM/DD") + "','" +  UOM.ToString() + "','" + QTY.ToString() + "','" +sP.ToString() + "','" + sPG.ToString() + "','" + MRP.ToString() + "','" + GST.ToString() + "','" + GSTamt.ToString() + "','" + CGST.ToString() + "','" + CGSTamt.ToString() + "','" + SGST.ToString() + "','" + SGSTamt.ToString() + "','" + UTGST.ToString() + "','" + UTGSTamt.ToString() + "','" + IGST.ToString() + "','" + IGSTamt.ToString() + "','"+sPT.ToString()+ "')", con);

                cmd.ExecuteNonQuery();
            }

        }
    }
}
