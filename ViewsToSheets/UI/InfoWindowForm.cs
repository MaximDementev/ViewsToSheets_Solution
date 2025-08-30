using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ViewsToSheets.UI
{
    public class InfoWindowForm : Form
    {
        private RichTextBox richTextBox;

        public InfoWindowForm(string text, IEnumerable<string> links)
        {
            InitializeComponent();
            InitializeEvents();
            PopulateContent(text, links);
        }

        private void InitializeComponent()
        {
            this.richTextBox = new RichTextBox();

            // Form settings
            this.Text = "Информация";
            this.Size = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;
            this.Padding = new Padding(10);

            // RichTextBox settings
            this.richTextBox.Dock = DockStyle.Fill;
            this.richTextBox.ReadOnly = true;
            this.richTextBox.BorderStyle = BorderStyle.None;
            this.richTextBox.BackColor = Color.WhiteSmoke;
            this.richTextBox.Font = new Font("Segoe UI", 10F);
            this.richTextBox.DetectUrls = true;
            
            // Add control
            this.Controls.Add(this.richTextBox);
        }

        private void InitializeEvents()
        {
            this.richTextBox.LinkClicked += RichTextBox_LinkClicked;
        }

        private void PopulateContent(string text, IEnumerable<string> links)
        {
            // Add main text
            richTextBox.Text = text;

            // Add links section if any exist
            bool hasLinks = false;
            foreach (var link in links)
            {
                if (!string.IsNullOrWhiteSpace(link))
                {
                    hasLinks = true;
                    break;
                }
            }

            if (hasLinks)
            {
                richTextBox.AppendText("\n\nСсылки:\n");
                foreach (string link in links)
                {
                    if (!string.IsNullOrWhiteSpace(link))
                    {
                        richTextBox.AppendText(link + "\n");
                    }
                }
            }
        }

        private void RichTextBox_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                // More reliable way to open URLs
                var psi = new ProcessStartInfo
                {
                    FileName = e.LinkText,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть ссылку: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
