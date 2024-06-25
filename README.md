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
using Spire.Doc;
using Spire.Pdf;



public partial class MainForm: Form
{
    private string file1Path;
    private string file2Path;

    private Button binSelectFile1;
    private Button binSelectFile2;
    private Button btnMerge;
    private Button btnSave;
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



    private void btnConvertToPDF_Click(object sender, EventArgs e)
    {
        try
        {
            // Load the documents
            Spire.Doc.Document doc1 = new Spire.Doc.Document();
            doc1.LoadFromFile(file1Path);
            Spire.Doc.Document doc2 = new Spire.Doc.Document();
            doc2.LoadFromFile(file2Path);

            // Convert to PDF
            string pdf1Path = Path.ChangeExtension(file1Path, ".pdf");
            string pdf2Path = Path.ChangeExtension(file2Path, ".pdf");
            doc1.SaveToFile(pdf1Path, Spire.Doc.FileFormat.PDF);
            doc2.SaveToFile(pdf2Path, Spire.Doc.FileFormat.PDF);

            // Now merge PDFs using Spire.Pdf
            string[] pdfFiles = new string[] { pdf1Path, pdf2Path };
            Spire.Pdf.PdfDocumentBase mergedPdf = Spire.Pdf.PdfDocument.MergeFiles(pdfFiles);

            // Save the merged PDF
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
            saveFileDialog.DefaultExt = "pdf";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Save merged PDF directly
                string mergedFilePath = saveFileDialog.FileName;
                File.WriteAllBytes(mergedFilePath, mergedPdf.SaveToBytes());
                MessageBox.Show("PDFs merged and saved successfully!");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to convert or merge PDFs: " + ex.Message);
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

    private void btnConvertToPDF_Click(object sender, EventArgs e)
    {
        try
        {
            // Load the DOCX documents
            Spire.Doc.Document doc1 = new Spire.Doc.Document();
            doc1.LoadFromFile(file1Path);
            Spire.Doc.Document doc2 = new Spire.Doc.Document();
            doc2.LoadFromFile(file2Path);

            // Convert to PDF
            string pdf1Path = Path.ChangeExtension(file1Path, ".pdf");
            string pdf2Path = Path.ChangeExtension(file2Path, ".pdf");
            doc1.SaveToFile(pdf1Path, Spire.Doc.FileFormat.PDF);
            doc2.SaveToFile(pdf2Path, Spire.Doc.FileFormat.PDF);

            // Merge PDFs using Spire.PDF
            Spire.Pdf.PdfDocumentBase mergedPdf = Spire.Pdf.PdfDocument.MergeFiles(new string[] { pdf1Path, pdf2Path });

            // Save the merged PDF
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf";
            saveFileDialog.DefaultExt = "pdf";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                mergedPdf.SaveToFile(saveFileDialog.FileName, Spire.Pdf.FileFormat.PDF);
                MessageBox.Show("PDFs merged and saved successfully!");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Failed to convert or merge PDFs: " + ex.Message);
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
        //btnConvertPDF
        this.btnConvertToPDF = new System.Windows.Forms.Button();
        // Setup the button properties
        this.btnConvertToPDF.Location = new System.Drawing.Point(50, 250);
        this.btnConvertToPDF.Size = new System.Drawing.Size(150, 23);
        this.btnConvertToPDF.Text = "Convert to and Merge PDF";
        this.btnConvertToPDF.Click += new EventHandler(this.btnConvertToPDF_Click);

       

        // Initialize TextBoxes for displaying file paths
        this.txtFile1Path.Location = new System.Drawing.Point(160, 50); // Adjust as needed
        this.txtFile1Path.Size = new System.Drawing.Size(300, 23); // Adjust as needed

        this.txtFile2Path.Location = new System.Drawing.Point(160, 100); // Adjust as needed
        this.txtFile2Path.Size = new System.Drawing.Size(300, 23); // Adjust as needed

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
        this.Controls.Add(this.btnConvertToPDF);

        //main
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 474);
            this.Controls.Add(this.binSelectFile1);
            this.Controls.Add(this.binSelectFile2);
            this.Controls.Add(this.btnMerge);
            this.Controls.Add(this.btnSave);
            this.Name = "mainform";
            this.Text = "DOCX Manager";
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
