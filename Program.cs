using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq.Expressions;
using iTextSharp.text;
using iTextSharp.text.pdf;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft;
using Microsoft.Office.Interop.Word;



public partial class MainForm : Form
{
    private string file1Path;
    private string file2Path;

    private Button binSelectFile1;
    private Button binSelectFile2;
    private Button btnMerge;
    private Button btnSave;
    private Button btnMergePDF;
    private Button btnSavePDF;
    private TextBox txtFile1Path;
    private TextBox txtFile2Path;

    public MainForm()
    {
        InitializeComponent();
    }

    private void btnSelectFile1_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            file1Path = openFileDialog.FileName;
            txtFile1Path.Text = file1Path; // Display selected file path
        }
    }

    private void btnSelectFile2_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            file2Path = openFileDialog.FileName;
            txtFile2Path.Text = file2Path; // Display selected file path
        }
    }

    private void btnMergePDF_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(file1Path) || string.IsNullOrEmpty(file2Path))
        {
            MessageBox.Show("Please select both files first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
        saveFileDialog.DefaultExt = "pdf";
        saveFileDialog.AddExtension = true;

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            string outputPdfPath = saveFileDialog.FileName;

            try
            {
                // Convert DOCX to PDF
                string pdf1 = ConvertDocxToPdf(file1Path);
                string pdf2 = ConvertDocxToPdf(file2Path);

                // Merge PDFs
                MergePdfs(new List<string> { pdf1, pdf2 }, outputPdfPath);

                // Clean up temporary PDF files
                File.Delete(pdf1);
                File.Delete(pdf2);

                MessageBox.Show("PDFs merged successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private string ConvertDocxToPdf(string docxPath)
    {
        string pdfPath = Path.ChangeExtension(docxPath, ".pdf");

        Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
        word.Visible = false;
        word.ScreenUpdating = false;

        try
        {
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(docxPath);
            doc.ExportAsFixedFormat(pdfPath, WdExportFormat.wdExportFormatPDF);
            doc.Close();
        }
        finally
        {
            word.Quit();
        }

        return pdfPath;
    }


    private void MergePdfs(List<string> pdfFiles, string outputFile)
    {
        using (FileStream stream = new FileStream(outputFile, FileMode.Create))
        using (iTextSharp.text.Document document = new iTextSharp.text.Document())
        using (PdfCopy pdf = new PdfCopy(document, stream))
        {
            document.Open();

            foreach (string file in pdfFiles)
            {
                using (PdfReader reader = new PdfReader(file))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = pdf.GetImportedPage(reader, i);
                        pdf.AddPage (page);
                    }
                }
            }
        }
    }


    private void btnMerge_Click(object sender, EventArgs e)
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Word Documents (*.docx)|*.docx";
        saveFileDialog.DefaultExt = "docx";
        saveFileDialog.AddExtension = true;

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                string mergedFilePath = saveFileDialog.FileName;
                File.Copy(file1Path, mergedFilePath, true);

                using (WordprocessingDocument mergedDoc = WordprocessingDocument.Open(mergedFilePath, true))
                {
                    MainDocumentPart mainPart = mergedDoc.MainDocumentPart;

                    using (WordprocessingDocument doc2 = WordprocessingDocument.Open(file2Path, false))
                    {
                        // Ensure that styles and definitions are copied or linked
                        CopyStyles(doc2, mergedDoc);

                        Body body2 = doc2.MainDocumentPart.Document.Body;
                        foreach (var element in body2.Elements())
                        {
                            mainPart.Document.Body.Append((OpenXmlElement)element.CloneNode(true));
                        }
                    }

                    mergedDoc.MainDocumentPart.Document.Save();
                }
                MessageBox.Show("Files merged successfully!");
            }
            catch (IOException ex)
            {
                MessageBox.Show($"An error occurred while accessing the files: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void CopyStyles(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
    {
        StyleDefinitionsPart sourceStylesPart = sourceDoc.MainDocumentPart.StyleDefinitionsPart;
        StyleDefinitionsPart targetStylesPart = targetDoc.MainDocumentPart.StyleDefinitionsPart;

        if (sourceStylesPart != null)
        {
            if (targetStylesPart == null)
            {
                // If there's no StyleDefinitionsPart in the target, we create it
                targetStylesPart = targetDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            }

            // Importing styles from source to target, ensuring any existing styles are not overwritten
            targetStylesPart.FeedData(sourceStylesPart.GetStream(FileMode.Open, FileAccess.Read));
        }
    }




    private void btnSave_Click(object sender, EventArgs e)
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            System.IO.File.Copy(file1Path, saveFileDialog.FileName, true);
            MessageBox.Show("File saved sucessfully!");
        }
    }

    private void InitializeComponent()
    {
        this.binSelectFile1 = new System.Windows.Forms.Button();
        this.binSelectFile2 = new System.Windows.Forms.Button();
        this.btnMerge = new System.Windows.Forms.Button();
        this.btnSave = new System.Windows.Forms.Button();
        this.txtFile1Path = new System.Windows.Forms.TextBox();
        this.txtFile2Path = new System.Windows.Forms.TextBox();
        this.btnMergePDF = new System.Windows.Forms.Button();
        this.btnMergePDF.Location = new System.Drawing.Point(160, 150);
        this.btnMergePDF.Size = new System.Drawing.Size(100, 23);
        this.btnMergePDF.Text = "Merge as PDF";
        this.btnMergePDF.Click += new EventHandler(this.btnMergePDF_Click);

        // Initialize TextBoxes for displaying file paths
        this.txtFile1Path.Location = new System.Drawing.Point(160, 50);
        this.txtFile1Path.Size = new System.Drawing.Size(300, 23);

        this.txtFile2Path.Location = new System.Drawing.Point(160, 100);
        this.txtFile2Path.Size = new System.Drawing.Size(300, 23);

        //btnSelectFile1
        this.binSelectFile1.Location = new System.Drawing.Point(50, 50); //Positoin
        this.binSelectFile1.Size = new System.Drawing.Size(100, 23);
        this.binSelectFile1.Text = "Select File 1";
        this.binSelectFile1.Click += new EventHandler(this.btnSelectFile1_Click);

        //btnSelectFile2
        this.binSelectFile2.Location = new System.Drawing.Point(50, 100); //Positoin
        this.binSelectFile2.Size = new System.Drawing.Size(100, 23);
        this.binSelectFile2.Text = "Select File 2";
        this.binSelectFile2.Click += new EventHandler(this.btnSelectFile2_Click);

        //btnMerge
        this.btnMerge.Location = new System.Drawing.Point(50, 150);
        this.btnMerge.Size = new System.Drawing.Size(100, 23);
        this.btnMerge.Text = "Merge files";
        this.btnMerge.Click += new EventHandler(this.btnMerge_Click);

        //btnSave
        this.btnSave.Location = new System.Drawing.Point(50, 200);
        this.btnSave.Size = new System.Drawing.Size(100, 23);
        this.btnSave.Text = "Save files";
        this.btnSave.Click += new EventHandler(this.btnSave_Click);


        // Add controls to the form
        this.Controls.Add(this.binSelectFile1);
        this.Controls.Add(this.binSelectFile2);
        this.Controls.Add(this.btnMerge);
        this.Controls.Add(this.btnSave);
        this.Controls.Add(this.txtFile1Path);
        this.Controls.Add(this.txtFile2Path);
        this.Controls.Add(this.btnMergePDF);


        //main
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(680, 474);
        this.Controls.Add(this.binSelectFile1);
        this.Controls.Add(this.binSelectFile2);
        this.Controls.Add(this.btnMerge);
        this.Controls.Add(this.btnSave);
        this.Name = "mainform";
        this.Text = "File Manager";
        this.ResumeLayout(false);


    }
}

namespace Combining_Docx
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new MainForm());
        }
    }
}
