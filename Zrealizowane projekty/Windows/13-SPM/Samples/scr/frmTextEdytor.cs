using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using FastColoredTextBoxNS;
using System.Globalization;
using OlimpComponents;
using System.ComponentModel;
using Aplikacja.FileUtils;

namespace EditFile
{
   public partial class frmTextEdytor : Form, ITextEdytor
   {
      private System.Text.Encoding _TextEncoding;
      //private string _pathText;
      private TypePhrase[] _listTypePhrase;
      private List<ReplaceItem> _listReplace = new List<ReplaceItem>();
      private BindingList<ReplaceItem> _blReplace = new BindingList<ReplaceItem>();
      private string _SymbolDir = string.Empty;

      // Style
      private TextStyle GreenStyle = new TextStyle(Brushes.Green, null, FontStyle.Italic);
      private TextStyle RedStyle = new TextStyle(Brushes.Red, null, FontStyle.Bold);
      private TextStyle BlueStyle = new TextStyle(Brushes.Blue, null, FontStyle.Italic);

      public event SaveExternalHandler SaveExternal;
      public event OpenPhraseExternalHandler OpenPhraseExternal;

      public frmTextEdytor()
      {
         InitializeComponent();
         textEditor.TextChanged += Edytor_TextChanged;
         dgv_ListReplace.DataSource = _blReplace;

         txt_FileInfo.Text = string.Empty;
         cms_OpenEmail.Tag = TypePhrase.Email;
         cms_OpenLink.Tag = TypePhrase.Link;
         cms_OpenDir.Tag = TypePhrase.Directory;
         cms_OpenFile.Tag = TypePhrase.File;
         cms_OpenTask.Tag = TypePhrase.Task;
      }

      public frmTextEdytor(bool showButtonClose) : this()
      {
         btn_Close.Visible = showButtonClose;
      }

      public frmTextEdytor(string initDirectory) : this()
      {
         openTextFileDialog.InitialDirectory = initDirectory;
         saveTextFileDialog.InitialDirectory = initDirectory;
      }

      public frmTextEdytor(bool showButtonClose, string initDirectory) : this(showButtonClose)
      {
         openTextFileDialog.InitialDirectory = initDirectory;
      }
      
      public string SymbolDir
      {
         get { return _SymbolDir; }
         set { 
                _SymbolDir = value;
                BuildMenuWstaw();
             }
      }

      private void BuildMenuWstaw()
      {
         try
         {
            if (cms_Insert.DropDownItems.Count > 0) cms_Insert.DropDownItems.Clear();

            if (Directory.Exists(_SymbolDir))
            {
               foreach (string filePath in Directory.GetFiles(_SymbolDir, "*.txt"))
               {
                  string txtName = Path.GetFileNameWithoutExtension(filePath);
                  ToolStripMenuItem newItem = new ToolStripMenuItem(txtName);
                  newItem.Click += (sender, e) =>
                  {
                     string symbolPath = Path.Combine(_SymbolDir, $"{newItem.Text}.txt");
                     if (File.Exists(symbolPath))
                     {
                        string symbol = File.ReadAllText(symbolPath);
                        symbol = symbol.Replace("{now::date}", $"{DateTime.Now.ToString("yyyy-MM-dd")}");
                        symbol = symbol.Replace("{now::time}", $"{DateTime.Now.ToString("HH:mm:ss")}");
                        textEditor.InsertText(symbol);
                     }
                  };
                  cms_Insert.DropDownItems.Add(newItem);
               }
            }
         }
         catch
         {
            throw;
         }
      }

      #region -(P)- Formatka ----------------------------------------------------------------------------------------------------
      private void Form_Load(object sender, EventArgs e)
      {
      }

      private void Form_Shown(object sender, EventArgs e)
      {
         StatControls();
      }

      private void Form_FormClosing(object sender, FormClosingEventArgs e)
      {
      }
      #endregion -(K)- Formatka

      #region -(P)- Metody publiczne --------------------------------------------------------------------------------------------
      public string TextEdit
      {
         get { return textEditor.Text; }
         set
         {
            textEditor.Text = value;
            textEditor.IsChanged = false;
            StatControls();
         }
      }

      public Form FormHandle => this;

      public string Title
      { 
         get { return Text; }
         set { Text = value; }
      }

      public TypePhrase[] OpenPhrase
      {
         set 
         { 
            _listTypePhrase = value;
         }
      }

      public void LoadFromFile(string filePath)
      {
         try
         {
            if (!File.Exists(filePath)) return;

            using (var reader = new System.IO.StreamReader(filePath, true))
            {
               _TextEncoding = reader.CurrentEncoding;
               textEditor.Invoke((MethodInvoker)(() => textEditor.Text = reader.ReadToEnd()));
            }
         }
         catch 
         {
            filePath = string.Empty;
         }
         finally
         {
            textEditor.IsChanged = false;
            textEditor.TextChanged += TextView_TextChanged;
            txt_FileInfo.Text = filePath;
            StatControls();
         }
      }

      private void TextView_TextChanged(object sender, TextChangedEventArgs e)
      {
         StatControls();
      }

      public void LoadFromBytes(byte[] textArray, string textPath)
      {
         try
         {
            using (var memory = new MemoryStream(textArray))
            {
               using (var reader = new System.IO.StreamReader(memory, true))
               {
                  _TextEncoding = reader.CurrentEncoding;
                  textEditor.Text = reader.ReadToEnd();
               }
            }
         }
         catch
         {
            textPath = null;
         }
         finally
         {
            textEditor.IsChanged = false;
            textEditor.TextChanged += TextView_TextChanged;
            txt_FileInfo.Text = textPath;
            StatControls();
         }
      }

      public void Clear()
      {
         textEditor.TextChanged -= TextView_TextChanged;
         _TextEncoding = System.Text.Encoding.Default;
         txt_FileInfo.Text = string.Empty;
         StatControls();
      }
      #endregion -(K)- Metody publiczne -----------------------------------------------------------------------------------------

      #region -(P)- Zdarzenia kontrolek -----------------------------------------------------------------------------------------

      // - Zamknij okno
      private void btn_Zamknij_Click(object sender, EventArgs e)
      {
         bool isClosed = false;
         TabPage tp = this.Parent as TabPage;
         if (tp != null)
         {
            ExtTabControl cm = tp.Parent as ExtTabControl;
            if (isClosed = (cm != null))
               cm.CloseForm(this);
         }
         if (!isClosed) Close();
      }

      #region -(P)- Zdarzenia kontrolek (formatka) -----------------------------------------------------------------------------

      // - zawijany tekst
      private void btn_ZawijaneWiersze_Click(object sender, EventArgs e)
      {
         btn_Wrap.Checked = !btn_Wrap.Checked;
         textEditor.WordWrap = btn_Wrap.Checked;
      }

      // - zbajdź
      private void btn_Znajdz_Click(object sender, EventArgs e)
      {
         textEditor.ShowFindDialog();
      }

      // - zamień
      private void btn_Zamien_Click(object sender, EventArgs e)
      {
         textEditor.ShowReplaceDialog();
      }

      // - pokaż / ukryj ustawienia listy
      private void btn_UstawieniaListy_Click(object sender, EventArgs e)
      {
         ToolList.Visible = !ToolList.Visible;
         btn_ParameterReplace.Enabled = !ToolList.Visible;
      }

      private void btn_ParameterReplace_Click(object sender, EventArgs e)
      {
         PanelReplace.Visible = !PanelReplace.Visible;
         btn_ParameterList.Enabled = !PanelReplace.Visible;
      }

      #endregion -(K)- Zdarzenia kontrolek (formatka) --------------------------------------------------------------------------

      #region -(P)- Zdarzenia kontrolek (edytor) -------------------------------------------------------------------------------
      #endregion -(K)- Zdarzenia kontrolek 

      #region -(P)- Zdarzenia kontrolek (menu kontekstowe edytora) -------------------------------------------------------------

      // - wstaw text z tag
      private void popup_WstawText(object sender, EventArgs e)
      {
         ToolStripItem tsi = sender as ToolStripItem;
         if (tsi == null) return;
         string txt = tsi.Tag.CastAsString();
         if (txt.IsNullOrEmpty()) return;

         textEditor.SelectionLength = 0;
         textEditor.SelectedText = txt;
      }

      // - kopiuj zaznaczony tekst do schowska
      private void cms_Copy_Click(object sender, EventArgs e)
      {
         string txt = textEditor.SelectedText.Trim();
         if (!txt.IsNullOrEmpty()) Clipboard.SetText(txt);
      }

      // - kopiuj zaznaczony tekst do schowska z formatowaniem
      private void cms_CopyFormat_Click(object sender, EventArgs e)
      {
         textEditor.Copy();
      }

      // - wklej ze schowka
      private void cms_Paste_Click(object sender, EventArgs e)
      {
         if (Clipboard.ContainsText()) textEditor.SelectedText = Clipboard.GetText();
      }

      #endregion -(K)- Zdarzenia kontrolek (menu kontekstowe edytora)

      #endregion -(K)- Zdarzenia kontrolek

      #region -(P)- Zdarzenia własne --------------------------------------------------------------------------------------------
      // -- Naciśnięcie klawisza
      private void Editor_KeyDown(object sender, KeyEventArgs e)
      {
         if (e.Modifiers == Keys.Control && e.KeyCode == Keys.V)
         {
            e.SuppressKeyPress = true;
            akcja_WklejZeSchowka();
         }
         else
         if (e.Modifiers == Keys.Control && e.KeyCode == Keys.C)
         {
            e.SuppressKeyPress = true;
            akcja_KopiujDoSchowka();
         }
         else
         if (e.Modifiers == Keys.Control && e.KeyCode == Keys.S)
         {
            e.SuppressKeyPress = true;
            btn_Save.PerformClick();
         }
      }

      // - Zmiana treści
      private void Edytor_TextChanged(object sender, FastColoredTextBoxNS.TextChangedEventArgs e)
      {
         int SumaKontrolka = textEditor.Text.GetHashCode();

         TaskSyntaxHighlight(e);
         StatControls();
      }

      // - naciśnięcie przycisku myszki
      private void Edytor_MouseDown(object sender, MouseEventArgs e)
      {
            //ipoup = popupEdytor.Items.Add("Wstaw datę");
            //ipoup.Tag = DateTime.Now.ToString("yyyy-MM-dd");
            //ipoup.Click += popup_WstawText;
      }

      private void Edytor_MouseMove(object sender, MouseEventArgs e)
      {
         Place p = textEditor.PointToPlace(e.Location);
         List<Style> ls = textEditor.GetStylesOfChar(p);
         if ((ls.IndexOf(BlueStyle) > -1) && (ModifierKeys.HasFlag(Keys.Control)))
         {
            textEditor.Cursor = Cursors.Hand;
            return;
         }
         textEditor.Cursor = Cursors.Default;
      }
      #endregion -(K)- Zdarzenia własne

      #region -(P)- Akcje -------------------------------------------------------------------------------------------------------
      private void akcja_KopiujDoSchowka()
      {
         if (!textEditor.SelectedText.IsNullOrEmpty()) textEditor.Copy();
      }

      private void akcja_WklejZeSchowka()
      {
         if (Clipboard.ContainsText()) textEditor.Paste();
      }

      private void cms_Cut_Click(object sender, EventArgs e)
      {
         textEditor.Cut();
      }

      private void cms_Delete_Click(object sender, EventArgs e)
      {
         textEditor.SelectedText = string.Empty;
      }

      private void cms_Open_Click(object sender, EventArgs e)
      {
         ToolStripMenuItem item = sender as ToolStripMenuItem;
         if (item == null) return;
         TypePhrase tp = TypePhrase.None;
         if (item.Tag is TypePhrase) tp = (TypePhrase)item.Tag;
         OpenPhraseExternal?.Invoke(textEditor.SelectedText, tp);
      }

      private void cms_ToolsDeleteEmptyLines_Click(object sender, EventArgs e)
      {
         SplitBlockLines sb = new SplitBlockLines(textEditor.SelectedText);
         sb.RemoveEmpty();
         textEditor.SelectedText = sb.Build(true);
      }

      private void cms_ToolsRevert_Click(object sender, EventArgs e)
      {
         SplitBlockWords sb = new SplitBlockWords(textEditor.SelectedText);
         sb.Reverse();
         textEditor.SelectedText = sb.Build();
      }

      private void FirstUpperCase_OnBeforeBuild(ref List<string> words)
      {
         for (int iii = 0; iii<words.Count; iii++) 
           words[iii] = words[iii].First().ToString(CultureInfo.InvariantCulture).ToUpperInvariant() + words[iii].Substring(1).ToLowerInvariant();
      }

      private void cms_ToolsFirstUpperCase_Click(object sender, EventArgs e)
      {
         SplitBlockWords sb = new SplitBlockWords(textEditor.SelectedText);
         sb.OnBeforeBuild += FirstUpperCase_OnBeforeBuild;
         textEditor.SelectedText = sb.Build();
      }

      private void cms_ToolsUpperCase_Click(object sender, EventArgs e)
      {
         textEditor.SelectedText = textEditor.SelectedText.ToUpperInvariant();
      }

      private void cms_ToolsLowerCase_Click(object sender, EventArgs e)
      {
         textEditor.SelectedText = textEditor.SelectedText.ToLowerInvariant();
      }

      private void cms_ToolsNoPL_Click(object sender, EventArgs e)
      {
         byte[] bytes = Encoding.GetEncoding(1251).GetBytes(textEditor.SelectedText);
         textEditor.SelectedText = Encoding.ASCII.GetString(bytes); ;
      }

      private void cms_ListCreate_Click(object sender, EventArgs e)
      {
         SplitBlockLines sb = new SplitBlockLines(textEditor.SelectedText);
         sb.RemoveEmpty();
         sb.AddSeparator(Direction.Right, txt_Separator.Text, txt_Prefix.Text, txt_Sufix.Text);
         textEditor.SelectedText = sb.Build();
      }

      private void FirstWord_OnBeforeBuild(ref List<string> lines)
      {
         for (int iii = 0; iii < lines.Count; iii++)
         {
            lines[iii] = lines[iii].Split(new[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
         }
      }

      private void cms_ListFirstWord_Click(object sender, EventArgs e)
      {
         SplitBlockLines sb = new SplitBlockLines(textEditor.SelectedText);
         sb.RemoveEmpty();
         sb.OnBeforeBuild += FirstWord_OnBeforeBuild;
         textEditor.SelectedText = sb.Build(true);
      }

      private void LastWord_OnBeforeBuild(ref List<string> lines)
      {
         for (int iii = 0; iii < lines.Count; iii++)
         {
            lines[iii] = lines[iii].Split(new[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
         }
      }

      private void cms_ListLastWord_Click(object sender, EventArgs e)
      {
         SplitBlockLines sb = new SplitBlockLines(textEditor.SelectedText);
         sb.RemoveEmpty();
         sb.OnBeforeBuild += LastWord_OnBeforeBuild;
         textEditor.SelectedText = sb.Build(true);
      }

      private void cms_ListAddPrefixSufix_Click(object sender, EventArgs e)
      {
         SplitBlockLines sb = new SplitBlockLines(textEditor.SelectedText);
         sb.AddSeparator(Direction.None, string.Empty, txt_Prefix.Text, txt_Sufix.Text);
         textEditor.SelectedText = sb.Build(true);
      }

      #endregion -(K)- Akcje ----------------------------------------------------------------------------------------------------

      #region -(P)- Stan kontrolek ----------------------------------------------------------------------------------------------
      private void StatControls()
      {
         bool isExternal = SaveExternal == null;
         ToolBox.BackColor = textEditor.IsChanged ? Color.LightCoral : SystemColors.Control;
         btn_Open.Visible = isExternal;
         btn_SaveAs.Visible = isExternal;
         btn_Save.Enabled = textEditor.IsChanged;
         btn_SaveAs.Enabled = true;
      }

      #endregion -(K)- Stan kontrolek -------------------------------------------------------------------------------------------

      #region -(P)- Sprawdzenie danych ------------------------------------------------------------------------------------------
      #endregion -(K)- Sprawdzenie danych ---------------------------------------------------------------------------------------

      #region -(P)- Rysowanie / obsługa siatek danych ---------------------------------------------------------------------------
      #endregion -(K)- Rysowanie / obsługa siatek danych ------------------------------------------------------------------------

      #region -(P)- Wspólne -----------------------------------------------------------------------------------------------------
      // - pobierz link
      private Boolean GetLink(String RegExpression, String Source, int CharPosition, ref String Link)
      {
         Regex rex = new Regex(RegExpression, RegexOptions.Compiled | RegexOptions.IgnoreCase);
         MatchCollection matches = rex.Matches(Source);
         foreach (Match match in matches)
         {
            if ((CharPosition >= match.Index) && (CharPosition <= match.Index + match.Length))
            {
               Link = match.Value;
               return true;
            }
         }
         return false;
      }

      private void TaskSyntaxHighlight(TextChangedEventArgs e)
      {
         e.ChangedRange.ClearStyle(GreenStyle, RedStyle, BlueStyle);
         e.ChangedRange.SetStyle(RedStyle, @"!!!.*$", RegexOptions.Multiline);
         e.ChangedRange.SetStyle(BlueStyle, @"---.*$", RegexOptions.Multiline);
         e.ChangedRange.SetStyle(GreenStyle, @"--.*$", RegexOptions.Multiline);
         e.ChangedRange.ClearFoldingMarkers();
      }


      #endregion -(K)- Wspólne --------------------------------------------------------------------------------------------------

      private void btn_Open_Click(object sender, EventArgs e)
      {
         if (openTextFileDialog.ShowDialog() == DialogResult.OK)
         {
            LoadFromFile(openTextFileDialog.FileName);
            txt_FileInfo.Text = openTextFileDialog.FileName;

            TabPage tp = this.Parent as TabPage;
            if (tp != null)
            {
               ExtTabControl cm = tp.Parent as ExtTabControl;
               cm.SetTabText(this, FileUtils.ShortName(txt_FileInfo.Text));
            }
         }
      }

      private void btn_Save_Click(object sender, EventArgs e)
      {
         bool result = false;
         if (SaveExternal == null)
         {
            if (File.Exists(txt_FileInfo.Text))
               result = action_Save(txt_FileInfo.Text);
            else
               result = action_SaveAs();
         }
         else
            SaveExternal?.Invoke(textEditor.Text, ref result);

         if (result)
         {
            if (SaveExternal == null) ChangeTabText(Path.GetFileNameWithoutExtension(txt_FileInfo.Text));

            textEditor.IsChanged = false;
            StatControls();
         }
      }

      private void btn_SaveAs_Click(object sender, EventArgs e)
      {
         action_SaveAs();
      }

      private void popupEdytor_Opening(object sender, System.ComponentModel.CancelEventArgs e)
      {
         bool isSelectedText = !textEditor.SelectedText.IsNullOrEmpty();
         bool isClipoardText = Clipboard.ContainsText();
         bool isOpenItems = isSelectedText && _listTypePhrase != null && _listTypePhrase.Length > 0;
         //
         cms_Copy.Enabled = isSelectedText;
         cms_CopyFormat.Enabled = isSelectedText;
         cms_Paste.Enabled = isClipoardText;
         cms_Cut.Enabled = isSelectedText;
         cms_Delete.Enabled = isSelectedText;
         // -
         cms_Open.Visible = isOpenItems;
         cms_OpenLink.Visible = isOpenItems && _listTypePhrase.Contains(TypePhrase.Link);
         cms_OpenEmail.Visible = isOpenItems && _listTypePhrase.Contains(TypePhrase.Email);
         cms_OpenDir.Visible = isOpenItems && _listTypePhrase.Contains(TypePhrase.Directory);
         cms_OpenFile.Visible = isOpenItems && _listTypePhrase.Contains(TypePhrase.File);
         cms_OpenTask.Visible = isOpenItems && _listTypePhrase.Contains(TypePhrase.Task);
         //
         cms_Insert.Enabled = ! string.IsNullOrEmpty(_SymbolDir);
         cms_Tools.Enabled = isSelectedText;
         cms_List.Enabled = isSelectedText;
      }

      private bool action_Save(string textPath)
      {
         bool result = false;
         try
         {
            File.WriteAllText(textPath, textEditor.Text);
            txt_FileInfo.Text = textPath;
            result = true;
         }
         catch (Exception ex)
         {
            MessageBox.Show($"Coś poszło nie tak ... {ex.Message}");
         }
         return result;
      }

      private bool action_SaveAs()
      {
         bool result = false;
         if (!txt_FileInfo.Text.IsNullOrEmpty())
         {
            // folder inicjalizacyjny (jeżeli istnieje)
            if (File.Exists(txt_FileInfo.Text))
            {
               saveTextFileDialog.InitialDirectory = Path.GetDirectoryName(txt_FileInfo.Text);
               saveTextFileDialog.FileName = Path.GetFileName(txt_FileInfo.Text);
            }
         }
         if (saveTextFileDialog.ShowDialog() == DialogResult.OK)
            result = action_Save(saveTextFileDialog.FileName);
         return result;
      }

      private void btnReplace_Clear_Click(object sender, EventArgs e)
      {
         dgv_ListReplace.DataSource = null;
         _listReplace.Clear();
         dgv_ListReplace.DataSource = _blReplace;
      }

      private string GetSeparator()
      {
         string result = string.Empty;
         if (rbSep_Equal.Checked) result = "=";
         else
         if (rbSep_Minus.Checked) result = "-";
         else
         if (rbSep_Semicolon.Checked) result = ";";
         else
         if (rbSep_Comma.Checked) result = ";";
         else
         if (rbSep_Tabulator.Checked) result = $"\t";
         return result;
      }

      private void btnReplace_PasteClipboard_Click(object sender, EventArgs e)
      {
         if (Clipboard.ContainsText())
         {
            string txt = Clipboard.GetText();
            string sep = GetSeparator();
            string[] lines = txt.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            _listReplace.Clear();
            foreach (string line in lines)
            {
               string textSearch = line.ArgumentInList(sep, 1).Trim();
               string textReplace = line.ArgumentInList(sep, 2).Trim();
               if (!textSearch.IsNullOrEmpty() && !textReplace.IsNullOrEmpty())
                  _listReplace.Add(new ReplaceItem() { TextSearch = textSearch, TextReplace = textReplace });
            }
            dgv_ListReplace.DataSource = _listReplace;
         }
         else
            MessageBox.Show("Schowek nie zawiera tekstu");
      }

      private void btnReplace_Go_Click(object sender, EventArgs e)
      {
         string txt = textEditor.Text;
         foreach (ReplaceItem item in _listReplace)
         {
            txt = txt.Replace(item.TextSearch, item.TextReplace);
         }
         textEditor.Text = txt;
         if (File.Exists(txt_FileInfo.Text))
         {
            saveTextFileDialog.InitialDirectory = Path.GetDirectoryName(txt_FileInfo.Text);
            txt_FileInfo.Text = string.Empty;
         }
         StatControls();
      }

      private void ChangeTabText(string title)
      {
         TabPage tp = this.Parent as TabPage;
         if (tp != null)
         {
            ExtTabControl cm = tp.Parent as ExtTabControl;
            if (cm == null) return;
            cm.SetTabText(tp, title);
         }

      }

   }

   internal partial class SplitBlockWords
   {
      private List<string> words = new List<string>();
      private List<string> separators = new List<string>();

      public delegate void BeforeBuild(ref List<string> words);
      public event BeforeBuild OnBeforeBuild;

      public SplitBlockWords(string txt)
      {
         string word = string.Empty;
         string separator = string.Empty;
         bool? black = null;
         foreach (char c in txt)
         {
            if (black == null) black = !char.IsWhiteSpace(c);
            if (!char.IsWhiteSpace(c))
            {
               if (black.Equals(false))
               {
                  black = true;
                  separators.Add(separator);
                  word = string.Empty;
               }
               word += c;
            }
            else
            {
               if (black.Equals(true))
               {
                  black = false;
                  words.Add(word);
                  separator = string.Empty;
               }
               separator += c;
            }
         }
         if (black.Equals(true))
            words.Add(word);
         else
            separators.Add(separator);

         if (separators.Count < words.Count) separators.Add(string.Empty);
      }

      public void Reverse()
      {
         words.Reverse();
      }

      public string Build()
      {
         OnBeforeBuild?.Invoke(ref words);

         string target = string.Empty;
         if (separators.Count > words.Count)
         {
            target = separators[0];
            separators.RemoveAt(0);
         }
         for (int iii = 0; iii < words.Count; iii++)
         {
            target += words[iii] + separators[iii];
         }
         return target;
      }
   }

   internal enum Direction
   {
      None,
      Left, 
      Right
   }

   internal partial class SplitBlockLines
   {
      private List<string> lines = new List<string>();

      public delegate void BeforeBuild(ref List<string> lines);
      public event BeforeBuild OnBeforeBuild;

      public SplitBlockLines(string txt)
      {
         lines.AddRange(txt.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
      }

      public void RemoveEmpty()
      {
         for (int iii = lines.Count-1; iii >= 0; iii--)
         {
            if (string.IsNullOrEmpty(lines[iii].Trim())) 
              lines.RemoveAt(iii);
            else
              lines[iii] = lines[iii].Trim();
         }
      }

      private string Symbol(string symbol)
      {
         if (symbol == null) symbol = string.Empty;
         symbol = symbol.Replace("#9", "\t");
         symbol = symbol.Replace("#10", "\r");
         symbol = symbol.Replace("#13", "\n");
         return symbol;
      }

      public void AddSeparator(Direction direct, string separator, string prefix = null, string sufix = null)
      {
         for (int iii = 0; iii < lines.Count; iii++)
         {
            if (direct == Direction.Left)
               lines[iii] = Symbol(separator) + lines[iii];
            else
            if (direct == Direction.Right)
            {
               lines[iii] = Symbol(prefix) + lines[iii] + Symbol(sufix);
               if (lines.Count - iii > 1) lines[iii] += Symbol(separator);
            }
            else
            if (direct == Direction.None)
               lines[iii] = Symbol(prefix) + lines[iii] + Symbol(sufix);
         }
      }

      public string Build(bool newLine = false)
      {
         OnBeforeBuild?.Invoke(ref lines);

         string target = string.Empty;
         for (int iii = 0; iii < lines.Count; iii++)
         {
            target += lines[iii];
            if (newLine && lines.Count - iii > 1) target += Environment.NewLine;
         }
         return target;
      }
   }

   internal partial class ReplaceItem
   {
      public string TextSearch { get; set; }
      public string TextReplace { get; set; }
   }
}
