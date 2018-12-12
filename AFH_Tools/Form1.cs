using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AFH_Tools
{

    public partial class AFH : Form
    {
        private const string FileName = "Notepad.exe";

        //string sharepoint_milli = @"https://4thtest.sharepoint.com/Operations/St%20Jude/Forms/AllItems.aspx?id=%2FOperations%2FSt%20Jude%2FMilli%20SOPs";
        //string sharepoint_hp = @"https://4thtest.sharepoint.com/:f:/g/Operations/Er0My5UqB75Gnu6pgG_LgFkBgqL8AAV_UBK9z-UPAnWPbg?e=6PD6UV";
        private string openair = @"https://www.openair.com/index.pl";
        private string operations = @"https://4thtest.sharepoint.com/Operations/SitePages/Home.aspx";
        private string servicenow = @"https://4thsource.service-now.com/nav_to.do?uri=%2Fhome.do";
        private string five9 = @"https://login.five9.com/";
        private string googlevoice = @"https://voice.google.com/messages";
        private string fxwell = @"https://fxwell.com/Account/Login?ReturnUrl=%2FPortal%2FHome%2FIndex";
        private string esiepa = @"https://accessps.express-scripts.com/epa/epa.html";
        private string yammer = @"https://www.yammer.com/4thsource.com/#/threads/company?type=general";
        private string staffhub = @"https://staffhub.office.com/app";
        private string link = "";

        //production links
        private string milli = @"Documents\Mill_pdf\";
        private string helpdesk = @"Documents\Helpdesk_pdf\";

        //string helpdesk = @"C:\Users\MichelleMaraj\Desktop\Helpdesk_pdf\";
        private string brierley = @"\\4thtest.sharepoint.com@SSL\Operations\Brierley\bp_sop\bp_sop_pdf\";
        private string notes_file = @"\\4thtest.sharepoint.com@SSL\Operations\St Jude\notes\";
        private string bp_notes_file = @"\\4thtest.sharepoint.com@SSL\Operations\Brierley\bp_notes\";
        private int counter = 0;
        private string getfile = "";

        public AFH()
        {
            InitializeComponent();
        }
        private static void linkMethod(string link)
        {
            try
            {
                System.Diagnostics.Process.Start(link);
            }
            catch (Exception ex)
            {

                MessageBox.Show("error" + ex);

            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = servicenow;
            linkMethod(link);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = openair;
            linkMethod(link);
        }

        private void linkLabel_sharepoint_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = operations;
            linkMethod(link);
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = five9;
            linkMethod(link);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = googlevoice;
            linkMethod(link);
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = fxwell;
            linkMethod(link);
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = esiepa;
            linkMethod(link);
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = yammer;
            linkMethod(link);
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            link = staffhub;
            linkMethod(link);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        //open dialog screen
        //@TODO need to change this to service now tab
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Form2 secondForm;
                secondForm = new Form2();
                secondForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("error" + ex);
            }
        }

        private string directories(string dir)
        {
            if (checkBox1.Checked)
            {
                dir = milli;
            }
            else if (checkBox2.Checked)
            {
                dir = helpdesk;
            }

            else if (tabControl_clients.SelectedIndex == 2)
            {
                dir = brierley;
            }


            return dir;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                listBox1.Items.Clear();
                richTextBox1.Clear();
                string[] files = Directory.GetFiles(milli);
                foreach (string file in files)
                {
                    listBox1.Items.Add(System.IO.Path.GetFileName(file));

                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(helpdesk);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(System.IO.Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
        }

        private void opensj_Click(object sender, EventArgs e)
        {
            try
            {
                String dir = "";
                axAcroPDF1.Visible = true;
                webBrowser1.Visible = false;
                dir = directories(dir);

                string curItem = listBox1.SelectedItem.ToString();
                //MessageBox.Show(curItem);
                //now to open file
                string new_file = dir + curItem;
                Directory.Exists(new_file);
                if (curItem.ToString().ToLower().Contains(".pdf"))
                {
                    axAcroPDF1.LoadFile(new_file);
                    //Html_file.Text = file;
                    //Uri uri = new Uri(new_file);
                    //webBrowser1.Navigate(uri);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("adobe reader not installed" + ex);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            String dir = "";
            dir = directories(dir);
            int track = 0;

            track = FindWordsInSops(dir, track);
        }

        private int FindWordsInSops(string dir, int track)
        {
            if (tabControl_clients.SelectedIndex == 2)
            {
                try
                {
                    listBox2.SelectedItems.Clear();
                    for (int i = listBox2.Items.Count - 1; i >= 0; i--)
                    {
                        using (PdfReader pdf = new PdfReader(dir + listBox2.Items[i].ToString().ToLower()))
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int p = 1; p <= pdf.NumberOfPages; p++)
                            {
                                sb.Append(PdfTextExtractor.GetTextFromPage(pdf, p));
                            }

                            var contents = sb.ToString();

                            if (contents.Contains(textBox2.Text.ToLower()))
                            {
                                listBox2.SetSelected(i, true);
                                track = 1;
                            }
                            else
                            {
                                track = 0;
                            }
                        }

                        if (listBox2.Items[i].ToString().ToLower().Contains(textBox2.Text.ToLower()) || track == 1)
                        {
                            listBox2.SetSelected(i, true);

                        }
                        else if (!listBox2.Items[i].ToString().ToLower().Contains(textBox2.Text.ToLower()))
                        {
                            listBox2.Items.RemoveAt(i);

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("cannot find file" + ex);
                }

            }
            else
            {
                try
                {
                    listBox1.SelectedItems.Clear();
                    for (int i = listBox1.Items.Count - 1; i >= 0; i--)
                    {
                        using (PdfReader pdf = new PdfReader(dir + listBox1.Items[i].ToString().ToLower()))
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int p = 1; p <= pdf.NumberOfPages; p++)
                            {
                                sb.Append(PdfTextExtractor.GetTextFromPage(pdf, p));
                            }

                            var contents = sb.ToString();

                            if (contents.Contains(textBox1.Text.ToLower()))
                            {
                                listBox1.SetSelected(i, true);
                                track = 1;
                            }
                            else
                            {
                                track = 0;
                            }
                        }

                        if (listBox1.Items[i].ToString().ToLower().Contains(textBox1.Text.ToLower()) || track == 1)
                        {
                            listBox1.SetSelected(i, true);

                        }
                        else if (!listBox1.Items[i].ToString().ToLower().Contains(textBox1.Text.ToLower()))
                        {
                            listBox1.Items.RemoveAt(i);

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("cannot find file" + ex);
                }
            }


            return track;
        }

        //refresh button
        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document WordDoc = new Microsoft.Office.Interop.Word.Document();
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            WordDoc.Close();

            axAcroPDF1.Visible = false;

            ap.Quit();
            webBrowser1.Visible = false;
            if (checkBox1.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(milli);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(System.IO.Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
            else if (checkBox2.Checked)
            {
                try
                {
                    listBox1.Items.Clear();
                    richTextBox1.Clear();
                    string[] files = Directory.GetFiles(helpdesk);
                    foreach (string file in files)
                    {
                        listBox1.Items.Add(System.IO.Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
        }

        //transfer files from docx to pdf and move to production
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Form3 thirdForm;
                thirdForm = new Form3();
                thirdForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("error" + ex);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(FileName);



        }

        private string NextFile(string path, ref int counter)
        {
            var filePath = "";
            if ((filePath = Directory.EnumerateFiles(path).Skip(counter).FirstOrDefault()) != null)
            {
                counter++;
            }
            else
            {
                counter = 0;
                filePath = Directory.EnumerateFiles(path).Skip(counter).FirstOrDefault();
                counter++;
            }
            return filePath;
        }



        private void button_forward_Click(object sender, EventArgs e)
        {

            getfile = NextFile(notes_file, ref counter);
            //MessageBox.Show(getfile);
            FileStream inFile = new FileStream(getfile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader reader = new StreamReader(inFile);
            string text = reader.ReadToEnd();
            richTextBox1.Text = text;
            reader.Close();
            inFile.Close();

        }

        private void button_back_Click(object sender, EventArgs e)
        {
            string getcurrentfile = System.IO.Path.GetFileName(getfile);
            System.Diagnostics.Process.Start(notes_file + getcurrentfile);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string getcurrentfile = System.IO.Path.GetFileName(getfile);
            MessageBox.Show(getcurrentfile + "will be deleted");
            File.Delete(notes_file + getcurrentfile);
        }

        private void brierleyRecapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //create and send bpweekly report
            System.Diagnostics.Process.Start("BrierleyWeeklyReport.exe");

        }
        //bp refresh button
        private void button13_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document WordDoc = new Microsoft.Office.Interop.Word.Document();
            Microsoft.Office.Interop.Word.Application ap = new Microsoft.Office.Interop.Word.Application();
            WordDoc.Close();

            axAcroPDF1.Visible = false;

            ap.Quit();
            webBrowser1.Visible = false;

            try
            {
                //clear list of files
                listBox2.Items.Clear();
                //clear note box
                richTextBox2.Clear();
                string[] files = Directory.GetFiles(brierley);
                foreach (string file in files)
                {
                    listBox1.Items.Add(System.IO.Path.GetFileName(file));

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("error" + ex);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            String dir = "";
            dir = directories(dir);
            int track = 0;

            track = FindWordsInSops(dir, track);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                String dir = "";
                axAcroPDF2.Visible = true;
                //webBrowser2.Visible = false;
                dir = directories(dir);

                string curItem = listBox2.SelectedItem.ToString();
                //MessageBox.Show(curItem);
                //now to open file
                string new_file = dir + curItem;

                if (curItem.ToString().ToLower().Contains(".pdf"))
                {
                    axAcroPDF2.LoadFile(new_file);
                    //Html_file.Text = file;
                    //Uri uri = new Uri(new_file);
                    //webBrowser1.Navigate(uri);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("adobe reader not installed" + ex);
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                try
                {
                    listBox2.Items.Clear();
                    richTextBox2.Clear();
                    string[] files = Directory.GetFiles(brierley);
                    foreach (string file in files)
                    {
                        listBox2.Items.Add(System.IO.Path.GetFileName(file));

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("error" + ex);
                }
            }
        }
        //bp forward
        private void button9_Click(object sender, EventArgs e)
        {
            getfile = NextFile(bp_notes_file, ref counter);
            //MessageBox.Show(getfile);
            FileStream inFile = new FileStream(getfile, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamReader reader = new StreamReader(inFile);
            string text = reader.ReadToEnd();
            richTextBox2.Text = text;
            reader.Close();
            inFile.Close();
        }
        //delete notes
        private void button7_Click(object sender, EventArgs e)
        {
            string getcurrentfile = System.IO.Path.GetFileName(getfile);
            MessageBox.Show(getcurrentfile + "will be deleted");
            File.Delete(bp_notes_file + getcurrentfile);
        }
        //create notes
        private void button8_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(FileName);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string getcurrentfile = System.IO.Path.GetFileName(getfile);
            System.Diagnostics.Process.Start(bp_notes_file + getcurrentfile);
        }

        private void helpdeskToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"Documents\Helpdesk\");
        }

        private void milliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"Documents\Milli\");
        }

        private void brierleyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(@"Documents\bp_sop\");
        }
    }
}
