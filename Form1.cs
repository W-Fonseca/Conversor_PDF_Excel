using System;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Excel = Microsoft.Office.Interop.Excel;
// feito por WELLINGTON JUVENAL FERREIRA FONSECA
namespace PDF_To_Excel {

    public partial class Janela : Form {
        Excel.Application xlApp = new
            Microsoft.Office.Interop.Excel.Application();
        public Janela() {
            InitializeComponent();

        }

        private void Janela_Load(object sender, EventArgs e) {


        }

        private void textBox1_TextChanged(object sender, EventArgs e) {
           

         
        }

        private void button1_Click(object sender, EventArgs e) {

            try {
                ConvertePDF pdftxt = new ConvertePDF();
                textBox2.Text = pdftxt.ExtrairTexto_PDF(textBox1.Text);

                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xls" }) {
                    if (sfd.ShowDialog() == DialogResult.OK) {
                        // Cria uma aplicação em Excel 
                        Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                        // Cria um novo WorkBook dentro do aplicativo Excel
                        Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                        // criando uma nova planilha com aba do excel na pasta de trabalho
                        Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                        // veja a folha de excel por trás do programa  
                        app.Visible = true;
                        // obtenha a referência da primeira folha. 
                        worksheet = workbook.Sheets[1];
                        // armazenar sua referência à planilha
                        worksheet = workbook.ActiveSheet;
                        //   
                        worksheet.Name = "plan1";
                        // armazenar parte do cabeçalho no Excel 
                        Clipboard.Clear();    // limpa valor na area de transferencia        
                        Clipboard.SetText(textBox2.Text);
                        worksheet.Paste();

                        // salve a aplicação
                        workbook.SaveAs(sfd.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        // saia da aplição
                        app.Quit();
                    }
                    

                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e) {

            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog()) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK) {
                    //obtenha o caminho do arquivo especificado
                    textBox1.Text = openFileDialog.FileName;

                    //Leia o conteúdo do arquivo em um fluxo
                    var fileStream = openFileDialog.OpenFile();


                }
            }

        }
        private void textBox2_TextChanged(object sender, EventArgs e) {

        }
    }
    public class ConvertePDF {
        public string ExtrairTexto_PDF(string caminho) {
            using (PdfReader leitor = new PdfReader(caminho)) {
                StringBuilder texto = new StringBuilder();
                for (int i = 1; i <= leitor.NumberOfPages; i++) {
                    texto.Append(PdfTextExtractor.GetTextFromPage(leitor, i));
                }
                return texto.ToString();
            }
        }
    }
}
