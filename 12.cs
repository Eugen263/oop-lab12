using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace VisitCardGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            string templateFilePath = "template.docx"; // шлях до файлу з шаблоном
            string outputFilePath = "output.docx"; // шлях до файлу з результатом

            // Отримуємо введені дані з форми
            string lastName = txtLastName.Text;
            string firstName = txtFirstName.Text;
            string companyName = txtCompanyName.Text;
            string phone = txtPhone.Text;
            string email = txtEmail.Text;

            // Створюємо список із даних для кожної візитки
            List<Dictionary<string, string>> data = new List<Dictionary<string, string>>();
            for (int i = 0; i < 10; i++)
            {
                Dictionary<string, string> cardData = new Dictionary<string, string>();
                cardData.Add("LastName", lastName);
                cardData.Add("FirstName", firstName);
                cardData.Add("CompanyName", companyName);
                cardData.Add("Phone", phone);
                cardData.Add("Email", email);
                data.Add(cardData);
            }

            // Заповнюємо шаблон даними та зберігаємо результат
            try
            {
                Application wordApp = new Application();
                Document wordDoc = wordApp.Documents.Open(templateFilePath);
                foreach (Dictionary<string, string> cardData in data)
                {
                    foreach (Field field in wordDoc.Fields)
                    {
                        if (field.Type == WdFieldType.wdFieldMergeField)
                        {
                            string fieldName = field.Code.Text.Replace("MERGEFIELD", "").Trim();
                            if (cardData.ContainsKey(fieldName))
                            {
                                field.Select();
                                wordApp.Selection.TypeText(cardData[fieldName]);
                            }
                        }
                    }
                    wordDoc.Sections.Last.Range.Select();
                    wordApp.Selection.InsertBreak(WdBreakType.wdPageBreak);
                }
                wordDoc.SaveAs(outputFilePath);
                wordDoc.Close();
                wordApp.Quit();
                MessageBox.Show("Візитки згенеровано успішно!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Сталася помилка при генерації візиток: " + ex.Message);
            }
        }
    }
}
