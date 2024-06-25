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
