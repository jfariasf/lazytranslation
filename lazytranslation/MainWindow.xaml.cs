using System;
using System.Collections.Generic;

using System.Windows;

using Word = Microsoft.Office.Interop.Word;

using System.ComponentModel;
using System.Windows.Forms;

namespace lazytranslation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string openFilePath;
        string saveFilePath;
        string sourceLanguage;
        string targetLanguage;
        Dictionary<String, String> languages = new Dictionary<String, String>()
        {
            {"Arabic","ar"},
            {"Bangla","bn"},
            {"Bulgarian","bg"},
            {"Chinese Simplified","zh-Hans"},
            {"Chinese Traditional","zh-Hant"},
            {"Croatian","hr"},
            {"Czech","cs"},
            {"Danish","da"},
            {"Dutch","nl"},
            {"English","en"},
            {"Estonian","et"},
            {"Finnish","fi"},
            {"French","fr"},
            {"German","de"},
            {"Greek","el"},
            {"Hebrew","he"},
            {"Hindi","hi"},
            {"Hungarian","hu"},
            {"Icelandic","is"},
            {"Italian","it"},
            {"Japanese","ja"},
            {"Korean","ko"},
            {"Latvian","lv"},
            {"Lithuanian","lt"},
            {"Norwegian","nb"},
            {"Polish","pl"},
            {"Portuguese","pt"},
            {"Romanian","ro"},
            {"Russian","ru"},
            {"Slovak","sk"},
            {"Slovenian","sl"},
            {"Spanish","es"},
            {"Swedish","sv"},
            {"Thai","th"},
            {"Turkish","tr"},
            {"Ukrainian","uk"},
            {"Vietnamese","vi"},
            {"Welsh","cy"}
         };
        

        public MainWindow()
        {
            InitializeComponent();
            System.Windows.Controls.ComboBox test =  sourceLanguageBox;
            browseFileButton.Click += MyButtonClick;
            processFileButton.Click += MyButtonClick;
            saveFileButton.Click += MyButtonClick;
            foreach (String language in languages.Keys) {
                sourceLanguageBox.Items.Add(language);
                targetLanguageBox.Items.Add(language);
            }
        }

        
      private void Worker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
          {
            BackgroundWorker worker = sender as BackgroundWorker;
            ReadDocument(sender);
            
         }

        void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = (e.ProgressPercentage);
            var value = (int)Math.Round((e.ProgressPercentage / progressBar1.Maximum) * 100);
            if (value > 100) { value = 100; }
            progressText.Content= value +"%";
            


        }
        void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

        }


        void MyButtonClick(object sender, EventArgs e)
        {
            System.Windows.Controls.Button button = sender as System.Windows.Controls.Button;
            if (sender == processFileButton)
            {
                //this.readDocument();
                var source = (sourceLanguageBox.SelectedItem as System.Windows.Controls.ListBoxItem);
                
                
                sourceLanguage = languages[sourceLanguageBox.SelectedItem.ToString()];
                targetLanguage = languages[targetLanguageBox.SelectedItem.ToString()];

                BackgroundWorker worker = new BackgroundWorker();
                worker.WorkerReportsProgress = true;
                worker.ProgressChanged += Worker_ProgressChanged;
                worker.DoWork += new DoWorkEventHandler( Worker_DoWork);
                worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
                worker.RunWorkerAsync();
            }
            else if (sender == browseFileButton) {
                // Displays an OpenFileDialog so the user can select a Cursor.  

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Word Files|*.docx";
                openFileDialog1.Title = "Select a Word File";
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Assign the cursor in the Stream to the Form's Cursor property.  
                    openFilePath = openFileDialog1.FileName;
                    sourcePathText.Text = openFilePath;
                }
            }
            else if (sender == saveFileButton)
            {
                // Displays an OpenFileDialog so the user can select a Cursor.  

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Word Files|*.docx";
                saveFileDialog1.Title = "Save a Word File";
                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Assign the cursor in the Stream to the Form's Cursor property.  
                    saveFilePath = saveFileDialog1.FileName;
                    targetPathText.Text = saveFilePath;
                }
            }
            //here you can check which button was clicked by the sender
        }



        public Word.Range TranslateText(Word.Range r, int par, String comment) {
            if (r.Text.Length > 2 && r.Text!="\r\n" && r.Text != "\r\n\t")
            {
                String result = TranslationEndPoint.TranslationAPI.translate(sourceLanguage,targetLanguage, r.Text).Result;
                try //dumb way
                {
                    var rangeFormat = r.ParagraphFormat.Duplicate;
                    if (result != "")
                    {
                        r.Text = result;
                        r.ParagraphFormat = rangeFormat;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Range Error " + r.Text + " -> " + result);
                }
            }
            return r;
        }
        public void SetProgressBarNonThread(int count) {
            this.Dispatcher.Invoke(() =>
            {
                progressBar1.Maximum = count;

            });
           }
        /*
         
             Useless Docx translation
             */
        public void ReadDocument(object sender) {

            List<Word.Range> TablesRanges = new List<Word.Range>();

            var wordApp = new Word.Application();
            if (openFilePath == null)
            {
                System.Windows.Forms.MessageBox.Show("Select a file first!");
                return;
            }
            else if (saveFilePath == null) {
                System.Windows.Forms.MessageBox.Show("Select a saving path!");
                return;
            }
            try
            {
                var doc = wordApp.Documents.OpenNoRepairDialog(FileName: @"" + openFilePath, ConfirmConversions: false, ReadOnly: false, AddToRecentFiles: false, NoEncodingDialog: true);
                for (int iCounter = 1; iCounter <= doc.Tables.Count; iCounter++)
                {
                    Word.Range TRange = doc.Tables[iCounter].Range;
                    TablesRanges.Add(TRange);
                }

                Boolean bInTable;
                SetProgressBarNonThread(doc.Paragraphs.Count);

                for (int par = 1; par <= doc.Paragraphs.Count; par++)
                {
                    
                    bInTable = false;
                    Word.Range r = doc.Paragraphs[par].Range;
                    foreach (Word.Range range in TablesRanges)
                    {
                        if (r.Start >= range.Start && r.Start <= range.End)
                        {
                            r = TranslateText(r, par, "Tables");
                            bInTable = true;
                            break;
                        }

                    }

                    if (!bInTable)
                    {
                        r = TranslateText(r, par, "Paragraph");
                    }

                    (sender as BackgroundWorker).ReportProgress(par);
                }
                doc.SaveAs2(saveFilePath);
                doc.Close();
                System.Windows.Forms.MessageBox.Show("The document has been translated.");
            }
            catch (System.Runtime.InteropServices.COMException e) {
                System.Windows.Forms.MessageBox.Show("Error opening the file.");
        

            }
            
        }
    }
}
