using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace nauchka
{
    public partial class Group : Form
    {
        string x;
        public Group(string y )
        {
            InitializeComponent();
            x = y;
        }

        private void Group_Load(object sender, EventArgs e)
        {
            label1.Text = x;
            DataSet group = new DataSet();
            group.ReadXml("http://timetable.sbmt.by/xml/group.xml");
            dataGridView1.DataSource = group.Tables[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows != null)
            {

                string groupfile = dataGridView1.SelectedRows[0].Cells["number"].Value.ToString();
                label2.Text = groupfile;
                Result r = new Result(label1.Text,label2.Text);
                r.Show();

            }

        }
    }
}
