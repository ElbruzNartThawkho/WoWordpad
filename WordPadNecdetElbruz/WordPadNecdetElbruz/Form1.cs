using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordPadNecdetElbruz;
using WordPadNecdetElbruz.Resources;

namespace WordPadNecdetElbruz
{
    public partial class Form1 : Form
    {
        private List<Document> _documents = new List<Document>();
        public Form1()
        {
            InitializeComponent();
            LoadFonts();
            LoadSizes();
            LoadStyles();
            MainPreparing();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void FontSeçme(object sender, EventArgs e)
        {

        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSettingsForText();
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Int32.Parse(toolStripComboBox2.Text);
            }
            catch (Exception exception)
            {
                return;
            }
            SetSettingsForText();
        }

        private void toolStripComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSettingsForText();
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            SetSettingsForText();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            ((Document)tabControl1.SelectedTab.Tag).IsEdit = true;
        }
        private void SetSettingsForText()
        {
            if (toolStripComboBox3.Items.Count > 0 && toolStripComboBox2.Items.Count > 0 && toolStripComboBox1.Items.Count > 0 && tabControl1.TabCount > 0)
            {
                ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionFont = new Font((string)toolStripComboBox3.SelectedItem, Single.Parse(toolStripComboBox2.Text), GetFontStyle());
            }
        }
        private FontStyle GetFontStyle()
        {
            FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), toolStripComboBox1.SelectedItem.ToString());
            return fontStyle;
        }
        private void seçiliSekmeyiKapatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(((Document)tabControl1.SelectedTab.Tag).IsEdit || ((Document)tabControl1.SelectedTab.Tag).IsEdit && ((Document)tabControl1.SelectedTab.Tag).PathDocument == null)
            {
                var res = MessageBox.Show("Değişiklikleri kayıt etmek ister misiniz?", "Kapanış Sekmesi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    SaveDoc();
                }
            }
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Dispose();
            tabControl1.SelectedTab.Dispose();
            tabControl1.SelectedIndex = tabControl1.TabPages.Count - 1;
            if (tabControl1.TabPages.Count == 0)
            {
                this.Close();
            }
        }

        private void kaydetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveDoc();
        }
        private void SaveDoc()
        {
            SaveFileDialog sf = new SaveFileDialog();
            sf.Title = "Dosya Kayıt";
            sf.Filter = "Zengin Metin Belgesi (*.rtf)|*|.rtf|";
            if (sf.ShowDialog() == DialogResult.OK)
                richTextBox1.SaveFile(sf.FileName, RichTextBoxStreamType.RichText);
            this.Text = sf.FileName;
        }
        private void LoadFonts()
        {
            var fontsCollection = new InstalledFontCollection();
            var ff = fontsCollection.Families;
            foreach (var item in ff)
            {
                toolStripComboBox3.Items.Add(item.Name);
            }
            toolStripComboBox3.SelectedIndex = 61;
        }
        private void LoadSizes()
        {
            toolStripComboBox2.Items.Add(8);
            toolStripComboBox2.Items.Add(9);
            toolStripComboBox2.Items.Add(10);
            toolStripComboBox2.Items.Add(11);
            toolStripComboBox2.Items.Add(12);
            toolStripComboBox2.Items.Add(14);
            toolStripComboBox2.Items.Add(16);
            toolStripComboBox2.Items.Add(18);
            toolStripComboBox2.Items.Add(20);
            toolStripComboBox2.Items.Add(22);
            toolStripComboBox2.Items.Add(24);
            toolStripComboBox2.Items.Add(26);
            toolStripComboBox2.Items.Add(28);
            toolStripComboBox2.Items.Add(36);
            toolStripComboBox2.Items.Add(48);
            toolStripComboBox2.Items.Add(72);
            toolStripComboBox2.SelectedIndex = 3;
        }
        private void LoadStyles()
        {
            toolStripComboBox1.Items.Add(FontStyle.Regular.ToString());
            toolStripComboBox1.Items.Add(FontStyle.Bold.ToString());
            toolStripComboBox1.Items.Add(FontStyle.Italic.ToString());
            toolStripComboBox1.Items.Add(FontStyle.Strikeout.ToString());
            toolStripComboBox1.Items.Add(FontStyle.Underline.ToString());
            toolStripComboBox1.SelectedIndex = 0;
        }
        private void MainPreparing()
        {
            var newDoc = new Document();
            tabControl1.SelectedTab.Tag = newDoc;
            _documents.Add(newDoc);
        }

        private void ekleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TabPage page = new TabPage("Yeni Döküman");
            var richTextBox = new RichTextBox();
            richTextBox.Location = new Point(3, 3);
            richTextBox.Dock = DockStyle.Fill;
            richTextBox.TextChanged += richTextBox1_TextChanged;
            richTextBox.SelectionChanged += richTextBox1_SelectionChanged;
            page.Controls.Add(richTextBox);
            tabControl1.TabPages.Add(page);
            tabControl1.SelectedTab = page;
            var newDoc = new Document();
            tabControl1.SelectedTab.Tag = newDoc;
            _documents.Add(newDoc);
        }

        private void richTextBox1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void yazdırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDialog1.ShowDialog();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            GetFontStyle();
        }

        private void toolStripComboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SetSettingsForText();
        }

        private void rToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionColor = colorDialog1.Color;
            }
        }

        private void rYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Left;
        }

        private void oToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Center;
        }

        private void lYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionAlignment = HorizontalAlignment.Right;
        }

        private void listele_CheckedChanged(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionBullet = listele.Checked;
        }

        private void açToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                foreach (var fileName in openFileDialog1.FileNames)
                {
                    OpenDoc(fileName);
                }
            }
        }
        private void OpenDoc(String fileName)
        {
            var conOk = true;
            foreach (TabPage tab in tabControl1.TabPages)
            {
                if (fileName == ((Document)tab.Tag).PathDocument)
                {
                    MessageBox.Show("Bu dosya zaten açık!", "Hata", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    tabControl1.SelectedTab = tab;
                    conOk = false;
                }
            }
            if (conOk)
            {
                var newDoc = new Document();
                newDoc.PathDocument = fileName;
                TabPage page = new TabPage(newDoc.GetNameDoc());
                var richTextBox = new RichTextBox();
                richTextBox.Location = new Point(3, 3);
                richTextBox.Dock = DockStyle.Fill;
                richTextBox.TextChanged += richTextBox1_TextChanged;
                richTextBox.EnableAutoDragDrop = true;
                richTextBox.SelectionChanged += richTextBox1_SelectionChanged;
                StreamReader sr = new StreamReader(fileName, Encoding.Default);
                string str = sr.ReadToEnd();
                if (Path.GetExtension(newDoc.PathDocument) == ".txt")
                {
                    richTextBox.Text = str;
                }
                else
                {
                    richTextBox.Rtf = str;
                }
                sr.Close();
                page.Controls.Add(richTextBox);
                tabControl1.TabPages.Add(page);
                tabControl1.SelectedTab = page;
                tabControl1.SelectedTab.Tag = newDoc;
                _documents.Add(newDoc);
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void kesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Cut();
        }

        private void kopyalaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Copy();
        }

        private void yapıştırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Paste();
        }

        private void silToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectedText = "";
        }

        private void tamamınıSeçToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectAll();
        }

        private void düzenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectedText = DateTime.Now.ToLongDateString();
        }

        private void uToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Undo();
        }

        private void rToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ((RichTextBox)tabControl1.SelectedTab.Controls[0]).Redo();
        }

        private void resimEkleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog2.ShowDialog() == DialogResult.OK)
            {
                ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionBackColor = colorDialog2.Color;
            }
        }

        private void arkaboyama_Click(object sender, EventArgs e)
        {
            if (colorDialog2.ShowDialog() == DialogResult.OK)
            {
                ((RichTextBox)tabControl1.SelectedTab.Controls[0]).SelectionBackColor = colorDialog2.Color;
            }
        }

        private void resToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Resim Dosyası|*.png;*.jpg;*.gif;*.jpeg;*.bmp";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using (var image = Image.FromFile(ofd.FileName))
                    {
                        Clipboard.SetImage(image);
                        richTextBox1.Paste();
                    }
                }
            }
        }
    }
}
