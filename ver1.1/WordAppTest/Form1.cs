using System;
using System.Printing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing.Printing;
using System.Threading;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {                
        FolderBrowserDialog FBD=new FolderBrowserDialog();

        public Form1()
        {
            InitializeComponent();                        
        }

        //public static class myPrinters
        //{
        //    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
        //    public static extern bool SetDefaultPrinter(string Name);
        //}
                                
        public void Format_otvet()
        {
            
            // Инициализация Word               
            this.Cursor = Cursors.WaitCursor;
            Word.Application o_Word = new Word.Application();
            Word.Document o_Doc = new Word.Document();
            //string defaultPrinter;
            //string n_PrintName = comboBox1.SelectedItem.ToString();
            //var printServer = new LocalPrintServer();
            //defaultPrinter = printServer.DefaultPrintQueue.Name;
            //myPrinters.SetDefaultPrinter(n_PrintName);
            string d_PrintName = o_Word.ActivePrinter;
            string n_PrintName = comboBox1.SelectedItem.ToString();
            if (chkVisible.Checked)
            {
                o_Word.Visible = false;
            }
            else
            {
                o_Word.Visible = true;
            }               
                        
            //Открытие и редактирование документа            
            Object filename = txtFrom.Text+"\\otvet.txt";            
            Object confirmConversions = Type.Missing;
            Object readOnly = Type.Missing;
            Object addToRecentFiles = Type.Missing;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = Type.Missing;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingOEMCyrillicII;
            Object visible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = Type.Missing;
                        
            if (!System.IO.File.Exists(filename.ToString()))
            {
                o_Word.Visible = true;
                MessageBox.Show("Файл "+filename.ToString()+" не найден.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
                                    
            o_Word.Documents.Open(ref filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);
            
            //Word.Document Doc = new Word.Document();
            o_Doc = o_Word.Documents.Application.ActiveDocument;
            o_Doc.PageSetup.LeftMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.RightMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.TopMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.BottomMargin = o_Word.CentimetersToPoints(1.27f);
            Word.Find o_Find = o_Word.Selection.Find;
            Word.Selection o_Sel = o_Word.Selection;

            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();
            o_Find.Text = "ФИО: Переходченко М.В.";
            //Object wrap = Type.Missing; ;
            //Object replace = Type.Missing; ;
            o_Find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: Type.Missing,
                Format: false,
                ReplaceWith: Type.Missing, Replace: Type.Missing);


            o_Sel.EndKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
            o_Sel.HomeKey(Unit: Word.WdUnits.wdStory, Extend: Word.WdMovementType.wdExtend);
            o_Sel.Delete();

            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();
            o_Find.Text = "----------------------------------------------------------------------------^p----------------------------------------------------------------------------";
            o_Find.Replacement.Text = "----------------------------------------------------------------------------^m----------------------------------------------------------------------------";
            Object wrap = Word.WdFindWrap.wdFindContinue;
            Object replace = Word.WdReplace.wdReplaceAll;
            o_Find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: Type.Missing, Replace: replace);
            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();
            int PagesCount = o_Word.Documents.Application.ActiveDocument.Content.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

            //Сохранение в формате .doc
            //Object s_fileName = txtFrom.Text + "\\arch\\otvet.doc";
            //Object fileformat = Word.WdSaveFormat.wdFormatDocument;
            //o_Doc.SaveAs(ref s_fileName, ref fileformat);

            o_Word.ActivePrinter = n_PrintName;
            Object background = Type.Missing;
            Object append = Type.Missing;
            Object range = Type.Missing;
            Object outputFileName = Type.Missing;
            Object from = Type.Missing;
            Object to = Type.Missing;
            Object item = Type.Missing;
            Object copies = Type.Missing;
            Object pages = Type.Missing;
            Object pageType = Type.Missing;
            Object printToFile = Type.Missing;
            Object collate = Type.Missing;
            Object fileName = Type.Missing;
            Object activePrinterMacGX = Type.Missing;
            Object manualDuplexPrint = Type.Missing;
            Object printZoomColumn = Type.Missing;
            Object printZoomRow = Type.Missing;
            Object printZoomPaperWidth = Type.Missing;
            Object printZoomPaperHeight = Type.Missing;
            o_Doc.PrintOut(ref background, ref append,
             ref range, ref outputFileName, ref from, ref to,
             ref item, ref copies, ref pages, ref pageType,
             ref printToFile, ref collate, ref activePrinterMacGX,
             ref manualDuplexPrint, ref printZoomColumn, ref printZoomRow,
             ref printZoomPaperWidth, ref printZoomPaperHeight);

            Thread.Sleep(6000);

            o_Word.ActivePrinter = d_PrintName;
            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            ((Word._Application)o_Word).Quit(ref saveChanges,
            ref originalFormat, ref routeDocument);
            o_Word = null;            
            //Process.Start(@"winword.exe", string.Format(@"{0} /mFilePrintDefault /mFileExit /q /n", s_fileName));
            //myPrinters.SetDefaultPrinter(defaultPrinter);
            this.Cursor = Cursors.Default;
            tSSlb.Text = "Файл " + filename + " отправлен на " + n_PrintName + " Всего страниц: " + PagesCount.ToString();               
            
        }
                
        public void Format_vedLic()
        { 
            // Инициализация Word
            this.Cursor = Cursors.WaitCursor;                     
            Word.Application o_Word = new Word.Application();
            Word.Document o_Doc = new Word.Document();
            if (chkVisible.Checked)
            {
                o_Word.Visible = false;
            }
            else
            {
                o_Word.Visible = true;
            }    
            string dt_pick = dTPicker.Value.ToString("dd.MM.yyyy");
            string dt_pick_save = dTPicker.Value.ToString("MMdd");
            string[] Filelst = System.IO.Directory.GetFiles(txtFrom.Text, "day*.txt");

            if (System.IO.File.Exists(Filelst.ToString()))
            {
                o_Word.Visible = true;
                MessageBox.Show("Файл " + Filelst.ToString() + " не найден.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }   
                        
            //Открытие и редактирование документа           
            //Object o_filename = txtFrom.Text + "\\DAY"+dt_pick_save+".TXT";
            Object o_filename = Filelst[0];
            Object confirmConversions = Type.Missing;
            Object readOnly = Type.Missing;
            Object addToRecentFiles = Type.Missing;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = Type.Missing;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingOEMCyrillicII;            
            Object visible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = Type.Missing;
            
            o_Word.Documents.Open(ref o_filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);

            //Редактирование документа
            o_Doc = o_Word.Documents.Application.ActiveDocument;
            o_Doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            o_Doc.PageSetup.LeftMargin = o_Word.CentimetersToPoints(1f);
            o_Doc.PageSetup.RightMargin = o_Word.CentimetersToPoints(1f);
            o_Doc.PageSetup.TopMargin = o_Word.CentimetersToPoints(1f);
            o_Doc.PageSetup.BottomMargin = o_Word.CentimetersToPoints(1f);
            Word.Find o_Find = o_Word.Selection.Find;
            Word.Selection o_Sel = o_Word.Selection;
                        
            o_Sel.WholeStory(); //выделяет весь текст
            o_Sel.Font.Size = 4;
            o_Sel.HomeKey(Unit:Word.WdUnits.wdStory,Extend:Word.WdMovementType.wdMove);
            //o_Sel.TypeText("Ведомость открытых лицевых счетов "+dt_pick+" г."+"\r\r");            
            o_Sel.TypeText("Ведомость открытых лицевых счетов " + dt_pick + " г.");
            o_Sel.TypeParagraph();
            o_Sel.TypeParagraph();
            o_Sel.EndKey(Unit: Word.WdUnits.wdLine, Extend: Word.WdMovementType.wdExtend);                        
            o_Sel.Delete();
            
            //Поиск и замена
            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();
            o_Find.Text = "^m";
            o_Find.Replacement.Text = " ";
            Object wrap = Word.WdFindWrap.wdFindContinue;
            Object replace = Word.WdReplace.wdReplaceAll;
            o_Find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: Type.Missing, Replace: replace);
            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();
            
            o_Sel.EndKey(Unit:Word.WdUnits.wdStory,Extend:Type.Missing);
            o_Sel.TypeParagraph();
            o_Sel.TypeParagraph();
            o_Sel.TypeText("Руководитель ________________ Алпеева Галина Ивановна");
            o_Sel.TypeParagraph();
            o_Sel.TypeParagraph();
            o_Sel.TypeText("Гл.бухгалтер ________________ Зотова Елена Сергеевна");
            
            //Сохранение в формате .doc
            Object s_fileName = txtFrom.Text + "\\arch\\DAY" + dt_pick_save + ".doc";
            Object fileformat = Word.WdSaveFormat.wdFormatDocument;            
            o_Doc.SaveAs(ref s_fileName,ref fileformat);
            
            //Выход из Word без сохранения
            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            ((Word._Application)o_Word).Quit(ref saveChanges,
            ref originalFormat, ref routeDocument);
            o_Word = null;
            this.Cursor = Cursors.Default;
            tSSlb.Text = "Файл " + o_filename + " сохранен в " + s_fileName;
            
        }

        public void Format_vipLS()
        {
            // Инициализация Word
            this.Cursor = Cursors.WaitCursor;            
            Word.Application o_Word = new Word.Application();
            Word.Document o_Doc = new Word.Document();            
            string d_PrintName = o_Word.ActivePrinter;
            string n_PrintName = comboBox1.SelectedItem.ToString();
            if (chkVisible.Checked)
            {
                o_Word.Visible = false;
            }
            else
            {
                o_Word.Visible = true;
            }    

            //Открытие и редактирование документа            
            Object filename = txtFrom.Text + "\\vips.txt";
            Object confirmConversions = Type.Missing;
            Object readOnly = Type.Missing;
            Object addToRecentFiles = Type.Missing;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = Type.Missing;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingOEMCyrillicII;
            Object visible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = Type.Missing;
            if (!System.IO.File.Exists(filename.ToString()))
            {
                o_Word.Visible = true;
                MessageBox.Show("Файл " + filename.ToString() + " не найден.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            o_Word.Documents.Open(ref filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);
                       
            o_Doc = o_Word.Documents.Application.ActiveDocument;
            //o_Doc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA3;
            o_Doc.PageSetup.PageWidth = o_Word.CentimetersToPoints(29.7f);
            o_Doc.PageSetup.PageHeight = o_Word.CentimetersToPoints(42f);
            o_Doc.PageSetup.LeftMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.RightMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.TopMargin = o_Word.CentimetersToPoints(1.27f);
            o_Doc.PageSetup.BottomMargin = o_Word.CentimetersToPoints(1.27f);
            Word.Find o_Find = o_Word.Selection.Find;
            Word.Selection o_Sel = o_Word.Selection;

            o_Sel.WholeStory();
            o_Sel.Font.Size = 9;

            o_Find.ClearFormatting();
            o_Find.Replacement.ClearFormatting();            
            Object p_findText = "Нереализованная отриц.";
            if (o_Find.Execute(ref p_findText,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: Type.Missing,
                Format: false,
                ReplaceWith: Type.Missing, Replace: Type.Missing))
            {
                o_Sel.EndKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
                o_Sel.HomeKey(Unit: Word.WdUnits.wdStory, Extend: Word.WdMovementType.wdExtend);
                o_Sel.Delete();
                o_Sel.Delete();
            }
            else
            {
                o_Find.ClearFormatting();
                o_Find.Replacement.ClearFormatting();            
                Object s_findText = "Расходы по процентам";
                if (o_Find.Execute(ref s_findText,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: Type.Missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: Type.Missing,
                Format: false,
                ReplaceWith: Type.Missing, Replace: Type.Missing))
                {
                    o_Sel.EndKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
                    o_Sel.HomeKey(Unit: Word.WdUnits.wdStory, Extend: Word.WdMovementType.wdExtend);
                    o_Sel.Delete();
                    o_Sel.Delete();
                }
                else
                {                    
                    MessageBox.Show("Текст поиска не найден.","Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    o_Word.Visible = true;
                    return; ;
                }
            }

            //int pars = o_Doc.Paragraphs.Count;// узнаем количество строк в документе


            int PagesCount = o_Word.Documents.Application.ActiveDocument.Content.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            for (int i = 1; i <= PagesCount; i++)
            {
                o_Sel.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, PagesCount, i);
                o_Sel.HomeKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
                o_Sel.MoveRight(Unit: Type.Missing, Count: 68, Extend: Word.WdMovementType.wdMove);
                o_Sel.ColumnSelectMode = true;
                o_Sel.MoveRight(Unit: Type.Missing, Count: 85, Extend: Word.WdMovementType.wdMove);
                o_Sel.MoveDown(Unit: Word.WdUnits.wdLine, Count: 109, Extend: Word.WdMovementType.wdMove);
                o_Sel.ColumnSelectMode = false;
                //MessageBox.Show("Остановка");
                o_Sel.Delete();                
                //o_Sel.HomeKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
                //o_Sel.MoveDown(Unit: Word.WdUnits.wdLine, Count: 109, Extend: Type.Missing);
            }

            o_Doc.PageSetup.PageWidth = o_Word.CentimetersToPoints(21f);
            o_Doc.PageSetup.PageHeight = o_Word.CentimetersToPoints(29.7f);
            //o_Doc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            o_Doc.PageSetup.LeftMargin = o_Word.CentimetersToPoints(2.35f);
            
            o_Word.ActivePrinter = n_PrintName;

            Object background = Type.Missing;
            Object append = Type.Missing;
            Object range = Type.Missing;
            Object outputFileName = Type.Missing;
            Object from = Type.Missing;
            Object to = Type.Missing;
            Object item = Type.Missing;
            Object copies = Type.Missing;
            Object pages = Type.Missing;
            Object pageType = Type.Missing;
            Object printToFile = Type.Missing;
            Object collate = Type.Missing;
            Object fileName = Type.Missing;
            Object activePrinterMacGX = Type.Missing;
            Object manualDuplexPrint = Type.Missing;
            Object printZoomColumn = Type.Missing;
            Object printZoomRow = Type.Missing;
            Object printZoomPaperWidth = Type.Missing;
            Object printZoomPaperHeight = Type.Missing;
            o_Doc.PrintOut(ref background, ref append,
             ref range, ref outputFileName, ref from, ref to,
             ref item, ref copies, ref pages, ref pageType,
             ref printToFile, ref collate, ref activePrinterMacGX,
             ref manualDuplexPrint, ref printZoomColumn, ref printZoomRow,
             ref printZoomPaperWidth, ref printZoomPaperHeight);

            o_Word.ActivePrinter = d_PrintName;

            Thread.Sleep(2000);

            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            ((Word._Application)o_Word).Quit(ref saveChanges,
            ref originalFormat, ref routeDocument);
            o_Word = null;

            this.Cursor = Cursors.Default;
            tSSlb.Text = "Файл " + filename + " отправлен на " + n_PrintName;
        }

        private void Format_vipS()
        {
            // Инициализация Word
            this.Cursor = Cursors.WaitCursor;            
            Word.Application o_Word = new Word.Application();
            Word.Document o_Doc = new Word.Document();
            string d_PrintName = o_Word.ActivePrinter;
            string n_PrintName = comboBox1.SelectedItem.ToString();
            if (chkVisible.Checked)
            {
                o_Word.Visible = false;
            }
            else
            {
                o_Word.Visible = true;
            }    
            string[] Filelst = System.IO.Directory.GetFiles(txtFrom.Text, "vp*.txt");
            
            if (System.IO.File.Exists(Filelst.Length.ToString()))
            {
                o_Word.Visible = true;
                MessageBox.Show("Файл " + Filelst.ToString() + " не найден.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //Открытие и редактирование документа
            Object filename = Filelst[0];
            Object confirmConversions = Type.Missing;
            Object readOnly = Type.Missing;
            Object addToRecentFiles = Type.Missing;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = Type.Missing;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingOEMCyrillicII;
            Object visible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = Type.Missing;
                        
            o_Word.Documents.Open(ref filename,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);

            o_Doc = o_Word.Documents.Application.ActiveDocument;
            o_Doc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA3;
            //o_Doc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape; // Word.WdOrientation.wdOrientPortrait
            //o_Doc.PageSetup.PageWidth = o_Word.CentimetersToPoints(29.7f);
            //o_Doc.PageSetup.PageHeight = o_Word.CentimetersToPoints(42f);            
            
            o_Word.ActivePrinter = n_PrintName;

            Object background = Type.Missing;
            Object append = Type.Missing;
            Object range = Type.Missing;
            Object outputFileName = Type.Missing;
            Object from = Type.Missing;
            Object to = Type.Missing;
            Object item = Type.Missing;
            Object copies = Type.Missing;
            Object pages = Type.Missing;
            Object pageType = Type.Missing;
            Object printToFile = Type.Missing;
            Object collate = Type.Missing;
            Object fileName = Type.Missing;
            Object activePrinterMacGX = Type.Missing;
            Object manualDuplexPrint = Type.Missing;
            Object printZoomColumn = Type.Missing;
            Object printZoomRow = Type.Missing;
            Object printZoomPaperWidth = Type.Missing;
            Object printZoomPaperHeight = Type.Missing;
            o_Doc.PrintOut(ref background, ref append,
             ref range, ref outputFileName, ref from, ref to,
             ref item, ref copies, ref pages, ref pageType,
             ref printToFile, ref collate, ref activePrinterMacGX,
             ref manualDuplexPrint, ref printZoomColumn, ref printZoomRow,
             ref printZoomPaperWidth, ref printZoomPaperHeight);

            o_Word.ActivePrinter = d_PrintName;

            Thread.Sleep(2000);

            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            ((Word._Application)o_Word).Quit(ref saveChanges,
            ref originalFormat, ref routeDocument);
            o_Word = null;

            this.Cursor = Cursors.Default;
            tSSlb.Text = "Файл " + filename + " отправлен на "+n_PrintName;
        }

        private void Print_Jornal()
        {
            // Инициализация Word
            this.Cursor = Cursors.WaitCursor;            
            Word.Application o_Word = new Word.Application();
            Word.Document o_Doc = new Word.Document();
            string d_PrintName = o_Word.ActivePrinter;
            string n_PrintName = comboBox1.SelectedItem.ToString();
            if (chkVisible.Checked)
            {
                o_Word.Visible = false;
            }
            else
            {
                o_Word.Visible = true;
            }    
            string path_toFile = txtJornal.Text + "\\" + comboBox2.SelectedItem.ToString();
            string dt_pick = dTPicker.Value.ToString("dd MMMM yyyy г.");

            object fileNameJorn = path_toFile;
            Object confirmConversions = Type.Missing;
            Object readOnly = Type.Missing;
            Object addToRecentFiles = Type.Missing;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = Type.Missing;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingOEMCyrillicII;
            Object visible = Type.Missing;
            Object openConflictDocument = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = Type.Missing;

            o_Word.Documents.Open(ref fileNameJorn,
                ref confirmConversions,
                ref readOnly,
                ref addToRecentFiles,
                ref passwordDocument,
                ref passwordTemplate,
                ref revert,
                ref writePasswordDocument,
                ref writePasswordTemplate,
                ref format,
                ref encoding,
                ref visible,
                ref openConflictDocument,
                ref openAndRepair,
                ref documentDirection,
                ref noEncodingDialog);
            o_Doc = o_Word.Documents.Application.ActiveDocument;

            Word.Selection o_Sel = o_Word.Selection;

            o_Sel.HomeKey(Unit: Word.WdUnits.wdStory, Extend: Word.WdMovementType.wdMove);
            o_Sel.EndKey(Unit: Word.WdUnits.wdLine, Extend: Type.Missing);
            o_Sel.TypeText(" " + dt_pick);

            o_Word.ActivePrinter = n_PrintName;
            Object background = Type.Missing;
            Object append = Type.Missing;
            Object range = Type.Missing;
            Object outputFileName = Type.Missing;
            Object from = Type.Missing;
            Object to = Type.Missing;
            Object item = Type.Missing;
            Object copies = Type.Missing;
            Object pages = Type.Missing;
            Object pageType = Type.Missing;
            Object printToFile = Type.Missing;
            Object collate = Type.Missing;
            Object fileName = Type.Missing;
            Object activePrinterMacGX = Type.Missing;
            Object manualDuplexPrint = Type.Missing;
            Object printZoomColumn = Type.Missing;
            Object printZoomRow = Type.Missing;
            Object printZoomPaperWidth = Type.Missing;
            Object printZoomPaperHeight = Type.Missing;
            o_Doc.PrintOut(ref background, ref append,
             ref range, ref outputFileName, ref from, ref to,
             ref item, ref copies, ref pages, ref pageType,
             ref printToFile, ref collate, ref activePrinterMacGX,
             ref manualDuplexPrint, ref printZoomColumn, ref printZoomRow,
             ref printZoomPaperWidth, ref printZoomPaperHeight);

            o_Word.ActivePrinter = d_PrintName;

            Thread.Sleep(2000);

            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Type.Missing;
            Object routeDocument = Type.Missing;
            ((Word._Application)o_Word).Quit(ref saveChanges,
            ref originalFormat, ref routeDocument);
            o_Word = null;

            this.Cursor = Cursors.Default;
            tSSlb.Text = "Файл " + comboBox2.SelectedItem.ToString() + " отправлен на " + n_PrintName;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegistryKey reg = Registry.CurrentUser.CreateSubKey("Software\\AppAutoPrint\\Settings");
            try { txtFrom.Text = reg.GetValue("path_From").ToString(); }
            catch { txtFrom.Text = ""; }
            try { txtJornal.Text = reg.GetValue("path_Jornal").ToString(); }
            catch { txtJornal.Text = ""; }

            PrinterSettings.StringCollection sc = PrinterSettings.InstalledPrinters;            
            for (int i = 0; i < sc.Count; i++)
            {
                comboBox1.Items.Add(sc[i]);
            }
            if (txtJornal.Text != " ")
            {
                string[] JornalList = System.IO.Directory.GetFiles(txtJornal.Text, "*.doc");                
                for (int j = 0; j < JornalList.Length; j++)
                {
                    comboBox2.Items.Add(new System.IO.FileInfo(JornalList[j]).Name);
                }
            }            

            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\AppAutoPrint\\NameOfFind");
            if (key == null)
            {
                RegistryKey keynew = Registry.CurrentUser.CreateSubKey("Software\\AppAutoPrint\\NameOfFind");
                keynew.SetValue("printer_Name", @"\\stavropolct\HP4350-OPERO");
                keynew.SetValue("jornal_Name", "Журнал_работы-Ежедневный.doc");
                int index = comboBox1.FindString(keynew.GetValue("printer_Name").ToString());
                int index2 = comboBox2.FindString(keynew.GetValue("jornal_Name").ToString());
                comboBox2.SelectedIndex = index2;
                comboBox1.SelectedIndex = index;
            }
            else
            {                
                int index = comboBox1.FindString(key.GetValue("printer_Name").ToString());
                int index2 = comboBox2.FindString(key.GetValue("jornal_Name").ToString());
                comboBox2.SelectedIndex = index2;
                comboBox1.SelectedIndex = index;
            }

            chkVisible.Checked=true;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            RegistryKey reg = Registry.CurrentUser.CreateSubKey("Software\\AppAutoPrint\\Settings");
            reg.SetValue("path_From", txtFrom.Text);            
            reg.SetValue("path_Jornal", txtJornal.Text);
        }                                       

        private void button1_Click(object sender, EventArgs e)
        {              
            Format_otvet();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Format_vedLic();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Format_vipS();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Format_vipLS();
        }
        
        private void txtFrom_DoubleClick(object sender, EventArgs e)
        {
            DialogResult result=FBD.ShowDialog();
            if (result==DialogResult.OK)
            {
                txtFrom.Text=FBD.SelectedPath;
            }

        }
                
        private void button5_Click(object sender, EventArgs e)
        {
            Print_Jornal();
        }

        private void txtJornal_DoubleClick(object sender, EventArgs e)
        {
            DialogResult result = FBD.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtJornal.Text = FBD.SelectedPath;
            }
        }
                                         

        }
}
