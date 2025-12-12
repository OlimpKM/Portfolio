using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using ShellApi.Open;
using ViewFile;
using GongSolutions.Shell;
using ContextMenu.MenuStripDirectory;
using Microsoft.VisualBasic;
using OlimpComponents;
using Aplikacja.FileUtils;
using MyLib.Security;

namespace FileManager
{
   public partial class frmFileManager : Form, IFileManager, IBeforeClose
   {
      private const string _KeyFormFind = "FormFindFiles";

      private string _InitDirectory;
      private string _TemplateDirectory;
      private string _CurrentDirectory;
      private string _PhraseKey;
      private bool _UserCloseForm;
      private object _ObjectLock = new Object();
      private List<Form> _ListView = new List<Form>();

      event FindFilesEventHandler OnFindFiles;
      public event DelegatBeforeClose BeforeClose;

      public frmFileManager()
      {
         InitializeComponent();
         _InitDirectory = null;
         _CurrentDirectory = null;
         _TemplateDirectory = null;
         lblFileName.Text = string.Empty;

         this.PreviewKeyDown += form_PreviewKeyDown;
         this.extTabManager.PreviewKeyDown += form_PreviewKeyDown;

         this.shellView.DoubleClick += new System.EventHandler(this.shellView_DoubleClick);
         this.shellView.Navigated += new System.EventHandler(this.shellView_Navigated);
         this.shellView.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(form_PreviewKeyDown);
         this.shellView.SelectionChanged += new System.EventHandler(ShellView_SelectionChanged);
         this.shellView.BeforeShowMenu += new ShellView.DelegateBeforeShowMenu(ShellView_BeforeShowMenu);

         this.btnBack.DropDownOpening += new System.EventHandler(this.btnBack_Popup);
         this.btnForward.DropDownOpening += new System.EventHandler(this.btnForward_Popup);

         // dyski lokalne
         foreach (string drive in GetAllLocalDrive())
         {
            ToolStripItem item = new ToolStripLabel(drive);
            item.Tag = drive;
            item.Click += new EventHandler(GotoFolder_Click);
            btnDrive.DropDownItems.Add(item);
         }
         // foldery specjanle
         btnFavorite.DropDownItems.Add(ItemToolStripFavorite("Pulpit", Environment.GetFolderPath(Environment.SpecialFolder.Desktop)));
         btnFavorite.DropDownItems.Add(ItemToolStripFavorite("Pobrane", System.Environment.ExpandEnvironmentVariables("%userprofile%/downloads/")));
         btnFavorite.DropDownItems.Add(ItemToolStripFavorite("Moje dokumenty", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)));
         btnFavorite.DropDownItems.Add(ItemToolStripFavorite("Moje obrazy", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)));
         btnFavorite.DropDownItems.Add(ItemToolStripFavorite("Moje muzyka", Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)));
      }

      public frmFileManager(string InitDirectory, bool UserCloseForm = true, string phraseKey = null) : this()
      {
         _InitDirectory = InitDirectory;
         _UserCloseForm = UserCloseForm;
         _PhraseKey = phraseKey;
      }

      event FindFilesEventHandler IFileManager.FindFilesEvent
      {
         add
         {
            lock (_ObjectLock) { OnFindFiles += value; }
         }
         remove
         {
            lock (_ObjectLock) { OnFindFiles -= value; }
         }
      }

      #region Form events
      // - odczytaj folder poczatkowy
      private void Form_Load(object sender, EventArgs e)
      {
         action_Navigate(_InitDirectory);
         StateControls();
      }

      // - ustawienie focus(a)
      private void form_Shown(object sender, EventArgs e)
      {
         if (this.CanFocus) this.Focus();
      }

      // podczas zamykania
      private void frm_FormClosed(object sender, FormClosedEventArgs e)
      {
         BeforeClose?.Invoke();
      }

      // - obsługa klawiatury
      private void form_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
      {
         if (e.KeyCode == Keys.F8)
         {
            e.IsInputKey = true;
            action_Delete();
         }
         if (e.KeyCode == Keys.F6)
         {
            e.IsInputKey = true;
            action_CopyName();
         }
         else
         if (e.Shift && e.KeyCode == Keys.F4)
         {
            e.IsInputKey = true;
            action_OpenAsFile();
         }
         else
         if (e.KeyCode == Keys.F4)
         {
            e.IsInputKey = true;
            action_OpenFile();
         }
         else
         if (e.KeyCode == Keys.F7)
         {
            e.IsInputKey = true;
            action_NewFolder();
         }
      }
      #endregion

      #region Public
      Form IFileManager.FormHandle => this;

      public string TemplateDir
      {
         get { return _TemplateDirectory; }
         set
         {
            _TemplateDirectory = value;
            StateControls();
         }
      }

      public void OpenFile(string filePath)
      {
         action_OpenFile(filePath);
      }

      public void OpenImageFile(string filePath, byte[] fileArray)
      {
         action_OpenImageFile(filePath, fileArray);
      }

      public void SelectDir(string pathName)
      {
         action_Navigate(pathName);
      }

      #endregion

      #region ShellView events
      // - nawigacja 
      private void shellView_Navigated(object sender, EventArgs e)
      {
         btnUp.Enabled = shellView.CanNavigateParent;
         btnBack.Enabled = shellView.CanNavigateBack;
         btnForward.Enabled = shellView.CanNavigateForward;
         ShowDirectoryInBar(shellView.CurrentFolder);
      }

      // - zmiana ilości zaznaczonych
      private void ShellView_SelectionChanged(object sender, EventArgs e)
      {
         int count = shellView.SelectedItems.Count();
         lblFileName.Text = count == 0 ? string.Empty : $"Zaznaczonych: {count}";
      }

      // - przed pokazaniem menu systemowe czy własne
      private void ShellView_BeforeShowMenu(object sender)
      {
         shellView.CustomContextMenuEnable = ModifierKeys.HasFlag(Keys.Control);
      }

      // - dwuklik myszy = otwarcie pliku 
      private void shellView_DoubleClick(object sender, EventArgs e)
      {
         foreach (var item in shellView.SelectedItems)
            action_OpenFile(item.FileSystemPath);
      }
      #endregion

      #region ShellView context menu
      private void cms_ShellView_Opening(object sender, System.ComponentModel.CancelEventArgs e)
      {
         bool isDir = false;
         bool isFile = false;
         string extFile = string.Empty;

         ShellItem item = shellView.SelectedItems.LastOrDefault();
         bool isSelected = item != null;
         if (isSelected)
         {
            isDir = item.IsFolder;
            isFile = item.IsFileSystem && File.Exists(item.FileSystemPath);
            if (!isDir) extFile = Path.GetExtension(item.ParsingName).ToLower();
         }

         cms_NewFolder.Visible = true;
         cms_OpenFile.Visible = isFile;
         cms_Security.Visible = isFile && !_PhraseKey.IsNullOrEmpty();
         cms_CopyName.Enabled = isSelected;
         cms_CopyPath.Enabled = isSelected;
      }

      private void cms_OpenFile_Click(object sender, EventArgs e)
      {
         foreach (var item in shellView.SelectedItems)
         {
            action_OpenFile(item.FileSystemPath);
         }
      }

      private void cms_NewFolder_Click(object sender, EventArgs e)
      {
         action_NewFolder();
      }

      private void cms_SecurityCrypt_Click(object sender, EventArgs e)
      {
         foreach (var file in GetSelectionItems())
         {
            XSec.CryptFile(file, _PhraseKey);
         }
      }

      private void cms_SecurityDecrypt_Click(object sender, EventArgs e)
      {
         foreach (var file in GetSelectionItems())
         {
            XSec.DecryptFile(file, _PhraseKey);
         }
      }

      private void cms_CopyName_Click(object sender, EventArgs e)
      {
         action_CopyName();
      }

      private void cms_CopyPath_Click(object sender, EventArgs e)
      {
         action_CopyPath();
      }

      #endregion

      #region Action
      // - odświeżenie danych (shellview)
      private void action_Refresh()
      {
         try
         {
            shellView.Refresh();
         }
         catch { }
      }

      // - przejdź do katalogu wyżej (shellview)
      private void action_NavigateParent()
      {
         try
         {
            shellView.NavigateParent();
         }
         catch
         {
            btnUp.Enabled = false;
         }
      }

      // - do katalogu z przodu (shellview)
      private void action_NavigateForward()
      {
         try
         {
            shellView.NavigateForward();
         }
         catch
         {
            btnForward.Enabled = false;
         }
      }

      // - do katalogu z tyłu (shellview)
      private void action_NavigateBack()
      {
         try
         {
            shellView.NavigateBack();
         }
         catch
         {
            btnBack.Enabled = false;
         }
      }

      // - przejdz do katalogu (shellview)
      private bool action_Navigate(string Folder)
      {
         bool result = false;
         try
         {
            shellView.CurrentFolder = new ShellItem(new Uri(Folder));
            shellView.RefreshContents();
            result = true;
         }
         catch { }
         return result;
      }

      // - usuń zaznaczone elementy (shellview)
      private bool action_Delete()
      {
         bool result = false;
         try
         {
            if (MessageBox.Show("Czy chcesz usunąć zaznaczone elementy:\n" + GetSelectionItemsText(), "Pytanie", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
               shellView.DeleteSelectedItems();
               shellView.RefreshContents();
            }
            result = true;
         }
         catch { }
         return result;
      }

      // - kopiuj do schowka nazwy zaznaczonych elementów
      private void action_CopyName()
      {
         try
         {
            string[] items = GetSelectionItems();
            for (int iii = 0; iii < items.Length; iii++)
               items[iii] = Path.GetFileName(items[iii]);

            Clipboard.SetText(string.Join(Environment.NewLine, items));
         }
         catch { }
      }

      // - kopiuj do schowka ścieżki zaznaczonych elementów
      private void action_CopyPath()
      {
         try
         {
            Clipboard.SetText(string.Join(Environment.NewLine, GetSelectionItems()));
         }
         catch { }
      }

      // - utworz pusty plik
      private void action_EmptyFile(ref string pathName)
      {
         if (Path.GetExtension(pathName).IsNullOrEmpty()) pathName += ".txt";
         File.WriteAllText(pathName, string.Empty);
      }

      // - otwórz plik (shellview)
      private void action_OpenAsFile()
      {
         string element = string.Empty;
         ShellItem item = shellView.SelectedItems.LastOrDefault();
         if (item != null && item.IsFileSystem) element = Path.GetFileName(item.FileSystemPath);

         string pathName = string.Empty;
         string result = Interaction.InputBox("Nazwa", "Otworz plik", element, Cursor.Position.X, Cursor.Position.Y);
         if (result.Length > 0)
         {
            pathName = Path.Combine(_CurrentDirectory, result);
            if (!File.Exists(pathName)) action_EmptyFile(ref pathName);
            ShellOpen.InvokeOpen(pathName);
         }
      }

      private void action_OpenFile(string pathFile)
      {
         if (pathFile.IsNullOrEmpty() || !File.Exists(pathFile)) return;

         bool isServed = false;
         if (ModifierKeys.HasFlag(Keys.Control))
         {
            string extFile = Path.GetExtension(pathFile).ToLower();
            string sKey = FileUtils.HashFilePath($"view::{pathFile}");

            isServed = extTabManager.TabExist(sKey);
            if (!isServed)
            {
               ViewFactory fact = new ViewFactory();
               IViewFile frm = fact.ViewFile(pathFile);
               if (frm != null)
               {
                  extTabManager.AddTab(frm.FormHandle, sKey, null, Color.Silver, $"Podgląd {FileUtils.ShortName(pathFile)}");
                  TabPage page = extTabManager.LocatePage(sKey);
                  if (page != null)
                  {
                     frm.LoadFromFile(pathFile);
                     frm.FormClosed += View_FormClosed;
                     page.ImageIndex = 2;
                     page.ToolTipText = pathFile;
                     isServed = true;
                     _ListView.Add(frm.FormHandle);
                     StateControls();
                  }
               }
            }
            else
               extTabManager.SelectTab_byKey(sKey);
         }
         if (!isServed) ShellOpen.InvokeOpen(pathFile);
      }

      private void View_FormClosed(object sender, FormClosedEventArgs e)
      {
         Form frm = sender as Form;
         if (frm != null)
         {
            _ListView.Remove(frm);
            StateControls();
         }
      }

      private void action_OpenImageFile(string pathFile, byte[] fileArray)
      {
         if (pathFile.IsNullOrEmpty()) return;

         string extFile = Path.GetExtension(pathFile).ToLower();
         string sKey = FileUtils.HashFilePath($"view::{pathFile}");

         bool isServed = extTabManager.TabExist(sKey);
         if (!isServed)
         {
            ViewFactory fact = new ViewFactory();
            IViewFile frm = fact.ViewFile(pathFile);
            if (frm != null)
            {
               extTabManager.AddTab(frm.FormHandle, sKey, null, Color.Silver, $"Podgląd {FileUtils.ShortName(pathFile)}");
               TabPage page = extTabManager.LocatePage(sKey);
               if (page != null)
               {
                  frm.LoadFromBytes(fileArray, pathFile);
                  frm.FormClosed += View_FormClosed;
                  page.ImageIndex = 2;
                  page.ToolTipText = pathFile;
                  isServed = true;
                  _ListView.Add(frm.FormHandle);
                  StateControls();
               }
            }
            else
              extTabManager.SelectTab_byKey(sKey);
         }
         if (!isServed) ShellOpen.InvokeOpen(pathFile);
      }

      private void action_OpenFile()
      {
         string pathName = string.Empty;
         ShellItem item = shellView.SelectedItems.LastOrDefault();
         if (item != null && item.IsFileSystem) pathName = item.FileSystemPath;

         action_OpenFile(pathName);
      }

      // - utworz nowy folder
      private void action_NewFolder()
      {
         string element = string.Empty;
         ShellItem item = shellView.SelectedItems.LastOrDefault();
         if (item != null) element = item.DisplayName;
         string result = Interaction.InputBox("Nazwa", "Nowy folder", element, Cursor.Position.X, Cursor.Position.Y);
         if (result.Length > 0)
         {
            string newDir = Path.Combine(_CurrentDirectory, result);
            if (!Directory.Exists(newDir))
            {
               Directory.CreateDirectory(newDir);
               shellView.RefreshContents();
            }
            else
               MessageBox.Show("Folder istnieje");
         }
      }

      #endregion

      #region Bar [toolStripHeat]
      // - zamknij formatkę
      private void closeButton_Click(object sender, EventArgs e)
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

      // - wyszukaj pliki i foldery
      private void findButton_Click(object sender, EventArgs e)
      {
         if (extTabManager.LocateForm(_KeyFormFind) == null)
         {
            IFindFiles frmFind = null;
            OnFindFiles?.Invoke(this, ref frmFind);
            if (frmFind != null)
            {
               frmFind.InitialDirectory = _CurrentDirectory;
               frmFind.OpenFileEvent += OpenFileEvent;
               frmFind.ChangeDirEvent += ChangeDirEvent;
               extTabManager.AddTab(frmFind.FormHandle, _KeyFormFind);
               TabPage page = extTabManager.LocatePage(_KeyFormFind);
               if (page != null)
               {
                  page.ImageIndex = 1;
                  page.ToolTipText = "Wyszukaj";
               }
            }
         }
         else
            extTabManager.SelectTab_byKey(_KeyFormFind);
      }

      // - zamknij podgląd
      private void closeViewButton_Click(object sender, EventArgs e)
      {
         for(int iii= _ListView.Count(); iii > 0; iii--)
           extTabManager.CloseForm(_ListView[iii-1]);
         StateControls();
      }


      private void ChangeDirEvent(object sender, string pathDir)
      {
         action_Navigate(pathDir);
         extTabManager.SelectedIndex = 0;
      }

      private void OpenFileEvent(object sender, string pathFile)
      {
         action_OpenFile(pathFile);
      }
      #endregion

      #region Bar [toolOperation]
      private void btnHome_Click(object sender, EventArgs e)
      {
         action_Navigate(_InitDirectory);
      }

      private void btnUp_Click(object sender, EventArgs e)
      {
         action_NavigateParent();
      }

      private void btnBack_ButtonClick(object sender, EventArgs e)
      {
         action_NavigateBack();
      }
      void btnBack_Popup(object sender, EventArgs e)
      {
         List<ToolStripItem> items = new List<ToolStripItem>();

         btnBack.DropDownItems.Clear();
         foreach (ShellItem f in shellView.History.HistoryBack)
         {
            ToolStripItem item = new ToolStripLabel(f.DisplayName);
            item.Tag = f;
            item.Click += new EventHandler(backButtonMenuItem_Click);
            items.Insert(0, item);
         }
         btnBack.DropDownItems.AddRange(items.ToArray());
      }
      private void backButtonMenuItem_Click(object sender, EventArgs e)
      {
         ToolStripItem item = sender as ToolStripItem;
         if (item == null) return;
         ShellItem folder = item.Tag as ShellItem;
         if (folder == null) return;
         shellView.NavigateBack(folder);
      }

      private void btnForward_ButtonClick(object sender, EventArgs e)
      {
         action_NavigateForward();
      }
      void btnForward_Popup(object sender, EventArgs e)
      {
         btnForward.DropDownItems.Clear();
         foreach (ShellItem f in shellView.History.HistoryForward)
         {
            ToolStripItem item = new ToolStripLabel(f.DisplayName);
            item.Tag = f;
            item.Click += new EventHandler(forwardButtonMenuItem_Click);
            btnForward.DropDownItems.Add(item);
         }
      }
      private void forwardButtonMenuItem_Click(object sender, EventArgs e)
      {
         ToolStripItem item = sender as ToolStripItem;
         if (item == null) return;
         ShellItem folder = item.Tag as ShellItem;
         if (folder == null) return;
         shellView.NavigateForward(folder);
      }

      private void btnRefresh_Click(object sender, EventArgs e)
      {
         action_Navigate(_CurrentDirectory);
      }

      private void GotoFolder_Click(object sender, EventArgs e)
      {
         ToolStripLabel label = sender as ToolStripLabel;
         if (label == null) return;
         string Folder = label.Tag as string;
         if (string.IsNullOrEmpty(Folder)) return;
         action_Navigate(Folder);
      }

      private void btnTemplates_Click(object sender, EventArgs e)
      {
         ContextMenuStripDirectory popup = new ContextMenuStripDirectory();
         popup.ClikedDirectory += Template_ClikedDirectory;
         popup.ClikedFile += Template_ClikedFile;
         popup.AddItemsFromDirectory(_TemplateDirectory);
         popup.Show(Cursor.Position);
      }
      private void Template_ClikedFile(object sender, ContextMenuStripDirectory cms, string pathfile)
      {
         try
         {
            string source = pathfile;
            string fileName = Path.GetFileNameWithoutExtension(source);
            string fileExt = Path.GetExtension(source);

            string target = Path.Combine(_CurrentDirectory, $"{fileName}{fileExt}");
            int i = 0;
            while (File.Exists(target))
            {
               i++;
               target = Path.Combine(_CurrentDirectory, $"{fileName}_{i}{fileExt}");
            }
            File.Copy(source, target);
         }
         catch { }
      }
      private void Template_ClikedDirectory(object sender, ContextMenuStripDirectory cms, string pathdir)
      {
         if (MessageBox.Show($"Czy skopiować pliki z folderu ?{Environment.NewLine}{pathdir}", "Pytanie", MessageBoxButtons.YesNo) == DialogResult.No) return;

         try
         {
            foreach (string pathfile in cms.GetFiles(pathdir))
            {
               string source = pathfile;
               string target = Path.Combine(_CurrentDirectory, Path.GetFileName(source));
               File.Copy(source, target);
            }
         }
         catch { }
      }

      #endregion

      #region Bar [toolDirectory]
      private void ShowDirectoryInBar(ShellItem shellpath)
      {
         _CurrentDirectory = null;
         toolDirectory.Items.Clear();
         if (shellpath.IsFileSystem)
         {
            _CurrentDirectory = shellpath.FileSystemPath;
            string[] dirs = _CurrentDirectory.Split(Path.DirectorySeparatorChar);
            string temporary = string.Empty;
            int iii = 0;
            foreach (var dir in dirs)
            {
               iii++;
               ToolStripItem element = new ToolStripLabel();
               element.Text = dir;
               element.Click += GotoFolder_Click;
               temporary = Path.Combine(temporary, dir + Path.DirectorySeparatorChar);
               element.Tag = temporary;
               toolDirectory.Items.Add(element);

               if (iii < dirs.Count())
               {
                  ToolStripItem separator = new ToolStripLabel();
                  separator.Text = Path.DirectorySeparatorChar.ToString();
                  toolDirectory.Items.Add(separator);
               }
            }
         }
      }
      #endregion

      #region Bar [toolPath]
      // - kopiój bieżącą ścieżkę do schowka
      private void btnCopyPath_Click(object sender, EventArgs e)
      {
         if (!string.IsNullOrEmpty(_CurrentDirectory))
         {
            string path = _CurrentDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()) 
                          ? _CurrentDirectory 
                          : _CurrentDirectory + Path.DirectorySeparatorChar.ToString();
            Clipboard.SetText(path);
         }
      }

      // - otwórz Windows Eksplorer
      private void btnExplorerWindows_Click(object sender, EventArgs e)
      {
         try
         {
            Process.Start(_CurrentDirectory);
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message.ToString(), "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
         }
      }

      // - otwórz shell command
      private void btnConsoleWindows_Click(object sender, EventArgs e)
      {
         Process p = new Process();
         p.StartInfo.FileName = @"cmd";
         p.StartInfo.WorkingDirectory = _CurrentDirectory;
         p.Start();
      }
      #endregion

      #region Controls state
      // - stan kontrolek
      private void StateControls()
      {
         closeButton.Visible = _UserCloseForm;
         closeViewButton.Visible = _ListView.Count > 0;
         findButton.Visible = OnFindFiles != null;
         btnHome.Visible = !_InitDirectory.IsNullOrEmpty();
         btnTemplates.Visible = !_TemplateDirectory.IsNullOrEmpty();
      }
      #endregion

      #region Utils
      private ToolStripItem ItemToolStripFavorite(string title, string pathFull)
      {
         ToolStripItem item = new ToolStripLabel(title);
         item.Tag = pathFull;
         item.Click += new EventHandler(GotoFolder_Click);
         return item;
      }
      private string[] GetAllLocalDrive()
      {
         List<string> drives = new List<string>();
         DriveInfo[] allDrives = DriveInfo.GetDrives();
         foreach (DriveInfo d in allDrives) drives.Add(d.Name);
         return drives.ToArray();
      }

      private string[] GetSelectionItems()
      {
         List<string> items = new List<string>();
         foreach(var item in shellView.SelectedItems)
            items.Add(item.FileSystemPath);
         return items.ToArray();
      }
      private string GetSelectionItemsText()
      {
         string result = string.Empty;
         foreach (string item in GetSelectionItems())
         {
            if (!result.IsNullOrEmpty()) result += Environment.NewLine;
            if (File.Exists(item))
               result += Path.GetFileName(item);
            else
            if (Directory.Exists(item))
               result += $"[KAT] {Path.GetFileName(item)}";
         }
         return result;
      }
      #endregion
   }
}
