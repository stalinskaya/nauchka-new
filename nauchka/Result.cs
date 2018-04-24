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
using System.Xml.Linq;
using System.Dynamic;
using Word = Microsoft.Office.Interop.Word;

namespace nauchka
{
    public partial class Result : Form
    {
        string n, f, courseNum;
        string lecturerName, speciality;
        public Result(string file, string num, string course,string lecName,string spec)
        {
            InitializeComponent();
            n = num;
            f = file;
            courseNum = course;
            lecturerName = lecName;
            speciality = spec;
        }
        DataTable lecturerData;
        private void CreateLecturerData()
        {
            lecturerData = new DataTable();
            lecturerData.Columns.Add("number", typeof(Int32));
            lecturerData.Columns.Add("topic", typeof(String));
            lecturerData.Columns.Add("type", typeof(String));
            lecturerData.Columns.Add("date", typeof(String));
            lecturerData.Columns.Add("time", typeof(String));
            lecturerData.Columns.Add("hoursAmount", typeof(Int32));
            lecturerData.Columns.Add("signature", typeof(String));
        }

        private void Result_Load(object sender, EventArgs e)
        {
            label1.Text = n;
            //label2.Text = courseNum;
            CreateLecturerData();

            dataGridView1.DataSource = lecturerData;
            dataGridView1.Columns["number"].HeaderText = "№";
            dataGridView1.Columns["number"].Width = 100;
            dataGridView1.Columns["topic"].HeaderText = "Тема";
            dataGridView1.Columns["topic"].Width = 250;
            dataGridView1.Columns["type"].HeaderText = "Тип";
            dataGridView1.Columns["type"].Width = 200;
            dataGridView1.Columns["date"].HeaderText = "Дата";
            dataGridView1.Columns["date"].Width = 150;
            dataGridView1.Columns["time"].HeaderText = "Время";
            dataGridView1.Columns["time"].Width = 150;
            dataGridView1.Columns["hoursAmount"].HeaderText = "Количество часов";
            dataGridView1.Columns["hoursAmount"].Width = 150;
            dataGridView1.Columns["signature"].HeaderText = "Роспись";
            dataGridView1.Columns["signature"].Width = 150;

            string path = "http://timetable.sbmt.by/shedule/lecturer/" + f;

            XDocument doc = XDocument.Load(path);
            int i = 1;
            var elemList =
                from el in doc.Descendants("lesson")
                where ((string)el.Element("group")).IndexOf(n) > -1
                select el;

            foreach (var elem in elemList)
            {
                DataRow tempRow = lecturerData.NewRow();
                tempRow["type"] = elem.Element("type").Value;
                tempRow["date"] = elem.Element("date").Value;
                tempRow["time"] = elem.Element("time").Value;
                tempRow["number"] = i;
                tempRow["hoursAmount"] = 2;
                i++;
                lecturerData.Rows.Add(tempRow);
            }

            string subject= dataGridView1.SelectedRows[0].Cells["subject"].Value.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        { 
            //Word.Application wa = new Word.Application();

            Word.Document wd = new Word.Document();
            wd.Activate();
            Object start = Type.Missing;
            Object end = Type.Missing;
            wd.Content.Font.Size = 12;
            wd.Content.Font.Name = "Times New Roman";
            wd.Content.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Word.Range rng = wd.Range(ref start, ref end);
            //rng.Select();
            Word.Paragraph wordparagraph = wd.Paragraphs.Add();
            wordparagraph.Range.Text = "Государственное учреждение образования \"Институт бизнеса и менеджмента технологий\" Белорусского государственного университета";
            wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Карточка №1";
            wordparagraph.Range.Font.Size = 14;
            wordparagraph.Range.Bold = 1;
            wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordparagraph.Range.InsertParagraphAfter();


            wordparagraph.Range.Text = "(учета проведенных занятий)";
            wordparagraph.Range.Font.Size = 14;
            wordparagraph.Range.Bold = 1;
            wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordparagraph.Range.InsertParagraphAfter();

            rng.SetRange(rng.End, rng.End);
            Object defaultTableBehavior = Type.Missing;
            Object autoFitBehavior = Type.Missing;
            Word.Table tbl = wd.Tables.Add(rng, 1, 7,ref defaultTableBehavior,ref autoFitBehavior );
            SetHeadings(tbl.Cell(1, 1), "№ п/п");
            SetHeadings(tbl.Cell(1, 2), "Тема");
            SetHeadings(tbl.Cell(1, 3), "Тип");
            SetHeadings(tbl.Cell(1, 4), "Дата");
            SetHeadings(tbl.Cell(1, 5), "Время занятий");
            SetHeadings(tbl.Cell(1, 6), "Количество часов");
            SetHeadings(tbl.Cell(1, 7), "Подпись");
            for (int i = 0; i < lecturerData.Rows.Count; i++)
            {
                Word.Row newRow = wd.Tables[1].Rows.Add(Type.Missing);
                newRow.Range.Font.Bold = 0;
                newRow.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                newRow.Cells[1].Range.Text = lecturerData.Rows[i][0].ToString();
                newRow.Cells[2].Range.Text = lecturerData.Rows[i][1].ToString();
                newRow.Cells[3].Range.Text = lecturerData.Rows[i][2].ToString();
                newRow.Cells[4].Range.Text = lecturerData.Rows[i][3].ToString(); ;
                newRow.Cells[5].Range.Text = lecturerData.Rows[i][4].ToString();
                newRow.Cells[6].Range.Text = lecturerData.Rows[i][5].ToString();
                newRow.Cells[7].Range.Text = lecturerData.Rows[i][6].ToString();
            }
        }
        static void SetHeadings(Word.Cell tblCell, string text)
        {
            tblCell.Range.Text = text;
            tblCell.Range.Borders.InsideColor = Word.WdColor.wdColorBlack;
            tblCell.Range.Borders.OutsideColor = Word.WdColor.wdColorBlack;
            tblCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }
    }
}

