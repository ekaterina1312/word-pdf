using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System.IO;

namespace WordToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //progressbar
                progressBar1.Value = 0;

                progressBar1.Maximum = 1;

                //объявляем ковертер
                DocToPDFConverter converter = new DocToPDFConverter();

                //объявляем openfiledialog и savefiledialog
                OpenFileDialog op = new OpenFileDialog();
                SaveFileDialog sv = new SaveFileDialog();

                sv.Filter = "Document |*.pdf";

                if (op.ShowDialog() == DialogResult.OK)
                {
                    if (Path.GetExtension(op.FileName) == ".docx")
                    {
                        //создаем документ docx
                        WordDocument doc = new WordDocument(op.FileName, FormatType.Docx);

                        //создаем документ pdf
                        PdfDocument pdf = converter.ConvertToPDF(doc);

                        progressBar1.Value++;

                        //освобждаем ресурсы конвертера
                        converter.Dispose();

                        //закрываем docx файл
                        doc.Close();

                        //сохранение pdf 
                        if (sv.ShowDialog() == DialogResult.OK)
                        {
                            pdf.Save(sv.FileName);

                        }

                        //закрытие pdf
                        pdf.Close();
                    }
                    else if (Path.GetExtension(op.FileName) == ".doc")
                    {
                        //создаем документ doc
                        WordDocument doc = new WordDocument(op.FileName, FormatType.Doc);

                        //создаем документ pdf
                        PdfDocument pdf = converter.ConvertToPDF(doc);

                        progressBar1.Value++;

                        //освобждаем ресурсы конвертера
                        converter.Dispose();

                        //закрываем doc файл
                        doc.Close();

                        //сохранение pdf 
                        if (sv.ShowDialog() == DialogResult.OK)
                        {
                            pdf.Save(sv.FileName);

                        }
                    }

                } 
            }
            catch (Exception ex)
            {
                MessageBox.Show("Пустой файл!");
            }

       }


        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                //folderbrowserdialog для выбора папки
                FolderBrowserDialog folderBrowser = new FolderBrowserDialog();

                // cоздаем конвертер
                DocToPDFConverter converter = new DocToPDFConverter();

                
              
                //массив пдф документов
                PdfDocument[] pdfs;

                if (folderBrowser.ShowDialog() == DialogResult.OK)
                {
                    //массив файлов из папки
                    string[] files = Directory.GetFiles(folderBrowser.SelectedPath);

                    progressBar1.Value = 0;

                    progressBar1.Maximum = files.Length;

                    //задаем размеры массива пдф файлов
                    pdfs = new PdfDocument[files.Length];
                    
                    //перебераем в цикле все файлы и конвертируем
                    for (int i = 0; i < files.Length; i++)
                    {

                        if (Path.GetExtension(files[i]) == ".docx")
                        {

                            pdfs[i] = converter.ConvertToPDF(new WordDocument(files[i], FormatType.Docx));

                            progressBar1.Value++;

                            pdfs[i].Save(System.IO.Path.GetDirectoryName(files[i]) + "/" + System.IO.Path.GetFileNameWithoutExtension(files[i]) + ".pdf");

                            pdfs[i].Close();

                        }
                        else if (Path.GetExtension(files[i]) == ".doc")
                        {
                            pdfs[i] = converter.ConvertToPDF(new WordDocument(files[i], FormatType.Docx));

                            progressBar1.Value++;
                            pdfs[i].Save(System.IO.Path.GetDirectoryName(files[i]) + "/" + System.IO.Path.GetFileNameWithoutExtension(files[i]) + ".pdf");

                            pdfs[i].Close();
                        }
                    }
                    //освобождаем ресурсы конвертера

                    converter.Dispose();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Пустой файл!");
            }
          }
        
    }
}
