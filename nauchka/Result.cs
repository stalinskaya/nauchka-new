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
using System.Text.RegularExpressions;

namespace nauchka
{
    public partial class Result : Form
    {
        string n, f;
        public Result(string file, string num)
        {
            InitializeComponent();
            n = num;
            f = file;
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
            CreateLecturerData();

            dataGridView1.DataSource = lecturerData;
            dataGridView1.Columns["number"].HeaderText = "№";
            dataGridView1.Columns["number"].Width = 20;
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word.Document wd = new Word.Document();
            wd.Activate();
            Object start = Type.Missing;
            Object end = Type.Missing;
            wd.Content.Font.Size = 12;
            wd.Content.Font.Name = "Times New Roman";
            wd.Content.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Word.Range rng = wd.Range(ref start, ref end);
            
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

            wordparagraph.Range.Text = "Преподаватель";
            wordparagraph.Range.Font.Size = 12;
            wordparagraph.Range.Bold = 0;
            wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Факультет";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Специальность";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Курс                                    \t\tГруппа             ";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Учебный год                                    \tСеместр             ";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Учебная дисциплина";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Количество часов по плану";
            wordparagraph.Range.InsertParagraphAfter();

            wordparagraph.Range.Text = "Форма занятий";
            wordparagraph.Range.InsertParagraphAfter();

            rng.SetRange(rng.End, rng.End);
            Object defaultTableBehavior = Type.Missing;
            Object autoFitBehavior = Type.Missing;
            Word.Table tbl = wd.Tables.Add(rng, 1, 7, ref defaultTableBehavior, ref autoFitBehavior);
            SetHeadings(tbl.Cell(1, 1), "№ п/п");
            SetHeadings(tbl.Cell(1, 2), "Тема");
            SetHeadings(tbl.Cell(1, 3), "Тип");
            SetHeadings(tbl.Cell(1, 4), "Дата");
            SetHeadings(tbl.Cell(1, 5), "Время");
            SetHeadings(tbl.Cell(1, 6), "Кол. часов");
            SetHeadings(tbl.Cell(1, 7), "Подпись");
            int i;
            for (i = 0; i < lecturerData.Rows.Count; i++)
            {
                string s = lecturerData.Rows[i][3].ToString();
                string[] w = s.Split('.');
                string res = w[1];
                if (res == curNumMonth)
                {
                    Word.Row newRow = wd.Tables[1].Rows.Add();
                    newRow.Range.Font.Bold = 0;
                    newRow.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    newRow.Cells[1].Range.Text = lecturerData.Rows[i][0].ToString();
                    tbl.Columns[1].SetWidth(27, Word.WdRulerStyle.wdAdjustSameWidth);
                    newRow.Cells[2].Range.Text = lecturerData.Rows[i][1].ToString();
                    tbl.Columns[2].SetWidth(170, Word.WdRulerStyle.wdAdjustSameWidth);
                    string typeOfLesson = lecturerData.Rows[i][2].ToString();
                    if (typeOfLesson == "Управляемая самостоятельная работа")
                    {
                        newRow.Cells[3].Range.Text = "УСР";
                    }
                    else if (typeOfLesson == "Лабораторная работа")
                    {
                        newRow.Cells[3].Range.Text = "Лаб";
                    }
                    else newRow.Cells[3].Range.Text = lecturerData.Rows[i][2].ToString();
                    tbl.Columns[3].SetWidth(60, Word.WdRulerStyle.wdAdjustSameWidth);
                    newRow.Cells[4].Range.Text = lecturerData.Rows[i][3].ToString();
                    tbl.Columns[4].SetWidth(65, Word.WdRulerStyle.wdAdjustSameWidth);
                    newRow.Cells[5].Range.Text = lecturerData.Rows[i][4].ToString();
                    tbl.Columns[5].SetWidth(44, Word.WdRulerStyle.wdAdjustSameWidth);
                    newRow.Cells[6].Range.Text = lecturerData.Rows[i][5].ToString();
                    tbl.Columns[6].SetWidth(40, Word.WdRulerStyle.wdAdjustSameWidth);
                    newRow.Cells[7].Range.Text = lecturerData.Rows[i][6].ToString();
                    tbl.Columns[7].SetWidth(60, Word.WdRulerStyle.wdAdjustSameWidth);
                }
            }
            //i++;
            //wd.Tables[1].Rows.Add();
            //tbl.Rows[i].Cells[0].Merge(tbl.Rows[i].Cells[2]);
        }
        private string curNumMonth;
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string curItem = listBox1.SelectedItem.ToString();
            if (curItem == "февраль")
            {
                curNumMonth = "02";
            }
            else if (curItem == "март")
            {
                curNumMonth = "03";
            }
            else if (curItem == "апрель")
            {
                curNumMonth = "04";
            }
            else if (curItem == "май")
            {
                curNumMonth = "05";
            }
            else if (curItem == "июнь")
            {
                curNumMonth = "06";
            }
            else if (curItem == "сентябрь")
            {
                curNumMonth = "09";
            }
            else if (curItem == "октябрь")
            {
                curNumMonth = "10";
            }
            else if (curItem == "ноябрь")
            {
                curNumMonth = "11";
            }
            else if (curItem == "декабрь")
            {
                curNumMonth = "12";
            }
        }

        static void SetHeadings(Word.Cell tblCell, string text)
        {
            tblCell.Range.Text = text;
            tblCell.Range.Borders.Enable = 1;
            tblCell.Range.ParagraphFormat.SpaceAfter = 0;
            tblCell.Range.ParagraphFormat.SpaceBefore = 0;
            tblCell.Range.Borders.InsideColor = Word.WdColor.wdColorBlack;
            tblCell.Range.Borders.OutsideColor = Word.WdColor.wdColorBlack;
            tblCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }
    }
}

