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

namespace nauchka
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows != null)
            {

                string prepodfile= dataGridView1.SelectedRows[0].Cells["filename"].Value.ToString();
                label1.Text = prepodfile;
                Group gr = new Group(label1.Text);
                gr.Show();
                
            }
            
            
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            DataSet xmldata = new DataSet();
            xmldata.ReadXml("http://timetable.sbmt.by/xml/lecturer.xml");
            dataGridView1.DataSource = xmldata.Tables[0];

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
