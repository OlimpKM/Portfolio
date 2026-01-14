using Aplikacja;
using Microsoft.Web.WebView2.Core;
using OlimpComponents;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EditFile
{
   public partial class frmHtmlEdytor : Form, ITextEdytor
   {
      private System.Windows.Forms.Timer _messageTimer;
      private bool _EdytorReady = false;
      private string _FilePath = string.Empty;
      private string _TinymcePath;

      public event SaveExternalHandler SaveExternal;
      public event OpenPhraseExternalHandler OpenPhraseExternal;

      public Form FormHandle => this;

      public string Title
      {
         get { return Text; }
         set { Text = value; }
      }

      public string TextEdit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
      public TypePhrase[] OpenPhrase { set => throw new NotImplementedException(); }

      public frmHtmlEdytor(string tinymcePath)
      {
         InitializeComponent();

         TabPage tp = this.Parent as TabPage;

         ucWebView.InitializationCompleted += ucWebView_InitializationCompleted;
         ucWebView.WebMessage += ucWebView_Message;
         ucWebView.WebError += ucWebView_Message;
         ucWebView.MessageReceived += ucWebView_MessageReceived;
         ucWebView.BeforeLoadPage += ucWebView_BeforeLoadPage;
         ucWebView.AfterLoadPage += ucWebView_AfterLoadPage;

         _TinymcePath = tinymcePath;
         LabelStatus.Text = string.Empty;
         ControlsStat();
      }

      #region (p) Formatka
      private void Form_Load(object sender, EventArgs e)
      { }

      private void Form_Shown(object sender, EventArgs e)
      {
         ControlsStat();
      }

      private void Form_FormClosing(object sender, FormClosingEventArgs e)
      {
         _EdytorReady = false;
         ucWebView.Dispose();
         ControlsStat();
      }
      #endregion (k) Formatka

      #region (p) Async metod
      private async Task SaveEditorContentAsync(string filePath, bool close = false)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         if (filePath.EndsWith(".html", StringComparison.OrdinalIgnoreCase))
            try
            {
               // Pobierz zawartość edytora z WebView2
               string content = await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync("getEditorContent()");

               // Oczyść zawartość (usuń dodatkowe cudzysłowy)
               string textContent = content.Trim('"');

               if (textContent != "null")
               {
                  // Upewnij się, że folder istnieje
                  Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                  // Zapisz plik w sposób asynchroniczny
                  using (var writer = new StreamWriter(filePath, false))
                  {
                     await writer.WriteAsync(textContent);
                  }

                  // zmień nazwę tab(a)
                  ChangeTabText(filePath);

                  ToolStripTop.BackColor = SystemColors.Control;
               }
            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message);
            }
         if (close) Close();
      }

      private async Task LoadEditorContentAsync(string filePath)
      {
         if (!ucWebView.IsBrowserReady) return;

         try
         {
            if (!File.Exists(filePath)) return;

            // odczytaj plik
            string fileContent = string.Empty;
            using (var reader = new StreamReader(filePath))
            {
               fileContent = await reader.ReadToEndAsync();
            }
            string scriptToSetContent = $"tinymce.get('editor').setContent(`{fileContent}`);";
            await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync(scriptToSetContent);
            // zmień nazwę tab(a)
            ChangeTabText(filePath);
         }
         catch
         {
            throw;
         }
      }
      #endregion (k) Async metod

      #region (p) Interface
      public async void LoadFromFile(string filePath)
      {
         _FilePath = filePath;
         if (_EdytorReady) await LoadEditorContentAsync(filePath);
      }

      public void LoadFromBytes(byte[] byteArray, string filePath)
      {
         throw new NotImplementedException();
      }

      public async void Clear()
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         string scriptToSetContent = $"tinymce.get('editor').setContent(``);";
         await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync(scriptToSetContent);
      }
      #endregion (k) Interface

      #region (p) Obsługa przycisków
      // zamknięcie edytora
      private void btn_Close_Click(object sender, EventArgs e)
      {
         TabPage tp = this.Parent as TabPage;

         bool isClosed = false;
         if (tp != null)
         {
            ExtTabControl cm = tp.Parent as ExtTabControl;
            if (isClosed = (cm != null))
               cm.CloseForm(this);
         }
         if (!isClosed) Close();
      }

      // otwórz plik
      private async void btn_Open_Click(object sender, EventArgs e)
      {
         string dir = string.Empty;
         if (!string.IsNullOrWhiteSpace(_FilePath))
         {
            if (Directory.Exists(_FilePath))
               dir = _FilePath;
            else
               dir = Path.GetDirectoryName(_FilePath);
         }
         if (string.IsNullOrEmpty(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

         OpenFileDialog ofd = new OpenFileDialog
         {
            Title = "Odczyt pliku",
            Filter = "Pliki html (*.html)|*.html;*.htm",
            InitialDirectory = dir
         };
         
         if (ofd.ShowDialog() == DialogResult.OK)
         {
            _FilePath = ofd.FileName;
            await LoadEditorContentAsync(_FilePath);
         }
      }

      // zapisz
      private async void btn_Save_Click(object sender, EventArgs e)
      {
         if (File.Exists(_FilePath ?? string.Empty))
            await SaveEditorContentAsync(_FilePath);
         else
            btn_SaveAs.PerformClick();
      }

      // zapisz jako
      private async void btn_SaveAs_Click(object sender, EventArgs e)
      {
         string dir = string.Empty;
         if (!string.IsNullOrWhiteSpace(_FilePath))
         {
            if (Directory.Exists(_FilePath))
               dir = _FilePath;
            else
               dir = Path.GetDirectoryName(_FilePath);
         }
         if (string.IsNullOrEmpty(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

         SaveFileDialog sfd = new SaveFileDialog
         {
            Title = "Zapis pliku",
            Filter = "Pliki html (*.html)|*.html;*.htm",
            DefaultExt = "html",
            InitialDirectory = dir
         };

         if (sfd.ShowDialog() == DialogResult.OK)
         {
            _FilePath = sfd.FileName;
            await SaveEditorContentAsync(_FilePath);
         }
      }

      // mała czcionka
      private async void FontSize_Small_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         string fontSize = "10px";
         await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync($"changeFontSize('{fontSize}')");
      }

      // normalna czcionka
      private async void FontSize_Normal_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         string fontSize = "14px";
         await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync($"changeFontSize('{fontSize}')");
      }

      // duża czcionka
      private async void FontSize_Big_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         string fontSize = "18px";
         await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync($"changeFontSize('{fontSize}')");
      }

      // wstaw obrazek z pliku
      private async void btn_ImageFromFile_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         var openFileDialog = new Microsoft.Win32.OpenFileDialog
         {
            Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp"
         };

         if (openFileDialog.ShowDialog() == true)
         {
            string filePath = openFileDialog.FileName;
            string base64Image = ConvertImageToBase64(filePath);

            // Wywołaj JavaScript, aby wstawić obraz Base64 do TinyMCE
            await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync($"insertImage('{base64Image}')");
         }
      }

      // wstaw obrazek ze schowka
      private async void btn_ImageFromClipboard_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         if (Clipboard.ContainsImage())
         {
            Image clipboardImage = Clipboard.GetImage();
            byte[] imageBytes = ConvertImageToJpeg(clipboardImage);
            string base64Image = ConvertImageToBase64(imageBytes, "JPG");
            await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync($"insertImage('{base64Image}')");
         }
         else
         {
            MessageBox.Show("Schowek nie zawiera obrazu.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
         }
      }

      // drukuj
      private async void btn_Print_Click(object sender, EventArgs e)
      {
         if (!(ucWebView.IsBrowserReady && _EdytorReady)) return;

         await ucWebView.WebBrowser.CoreWebView2.ExecuteScriptAsync("printEditorContent();");
      }

      #endregion (k) Obsługa przycisków

      #region (p) Stan kontrolek
      private void ControlsStat()
      {
         btn_Open.Enabled = _EdytorReady;
         btn_Save.Enabled = _EdytorReady;
         btn_SaveAs.Enabled = _EdytorReady;
         btn_FontSize.Enabled = _EdytorReady;
         btn_Image.Enabled = _EdytorReady;
         btn_Print.Enabled = _EdytorReady;
      }
      #endregion (k) Stan kontrolek

      #region (p) Zdarzenia
      private void ucWebView_InitializationCompleted(object obj)
      {
         ControlsStat();
         string tinymcePath = Path.GetFullPath(_TinymcePath).Replace("\\", "/");
         ucWebView.LoadPage($"file:///{tinymcePath}");
      }

      private void ucWebView_Message(object sender, MessageEventArgs e)
      {
         StatusMessgeWithTime(e.Message);
      }

      private async void ucWebView_MessageReceived(object sender, MessageEventArgs e)
      {
         if (e.Message == "editorReady")
         {
            _EdytorReady = true;
            // Załadowanie pliku dokumentu
            if (File.Exists(_FilePath)) await LoadEditorContentAsync(_FilePath);
            ControlsStat();
         }
         else
         if (e.Message == "editorChange")
         {
            ToolStripTop.BackColor = Color.LightCoral;
            ControlsStat();
         }
         else
         if (e.Message == "editorSave")
         {
            btn_Save.PerformClick();
         }
      }

      private void ucWebView_BeforeLoadPage(object sender, BeforeLoadEventArgs e)
      {
         ControlsStat();
      }

      private void ucWebView_AfterLoadPage(object sender, AfterLoadEventArgs e)
      {
         ControlsStat();
      }
      #endregion (k) Zdarzenia

      #region (p) Komunikaty
      // wyświetl komunikat przez określony czas
      private void StatusMessgeWithTime(string message, int durationMs = 3000)
      {
         LabelStatus.Text = message;

         if (_messageTimer != null)
         {
            _messageTimer.Stop();
            _messageTimer.Dispose();
         }

         _messageTimer = new System.Windows.Forms.Timer
         {
            Interval = durationMs
         };

         // Dodajemy zdarzenie Tick, które zostanie wywołane po upływie czasu
         _messageTimer.Tick += (sender, e) =>
         {
            var timer = sender as System.Windows.Forms.Timer;
            if (timer != null)
            {
               LabelStatus.Text = string.Empty;
               timer.Stop();
               timer.Dispose();
            }
         };
         _messageTimer.Start();
      }
      #endregion (k) Komunikaty

      #region (p) Pozostałe
      private string ConvertImageToBase64(string filePath)
      {
         byte[] imageBytes = File.ReadAllBytes(filePath);
         return $"data:image/{Path.GetExtension(filePath).TrimStart('.')};base64,{Convert.ToBase64String(imageBytes)}";
      }

      private string ConvertImageToBase64(byte[] imageBytes, string fileExt)
      {
         return $"data:image/{fileExt.TrimStart('.')};base64,{Convert.ToBase64String(imageBytes)}";
      }

      private byte[] ConvertImageToJpeg(Image image)
      {
         using (MemoryStream ms = new MemoryStream())
         {
            // Zapisujemy obraz do MemoryStream w formacie JPG
            image.Save(ms, ImageFormat.Jpeg);
            return ms.ToArray(); // Zwracamy tablicę bajtów
         }
      }

      private void ChangeTabText(string filePath)
      {
         // zmień nazwę tab(a)
         TabPage tp = this.Parent as TabPage;
         if (tp != null)
         {
            ExtTabControl cm = tp.Parent as ExtTabControl;
            if (cm != null) cm.SetTabText(tp, Path.GetFileNameWithoutExtension(filePath));
         }
      }
      #endregion (p) Pozostałe
   }
}
