using GoggleDriveApp;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

public class GoogleDriveHelper
{
   private readonly string _clientSecretFile;
   private readonly string _tokenStorePath;
   private DriveService _service;
   private Form _owner;

   public GoogleDriveHelper(string clientSecretFile, string tokenStorePath)
   {
      _clientSecretFile = clientSecretFile;
      _tokenStorePath = tokenStorePath;
   }

   public void SetOwner(Form owner = null)
   {
      _owner = owner;
   }

   // - Autoryzacja
   public async Task<DriveService> GetServiceAsync()
   {
      if (_service != null)
         return _service;

      UserCredential credential;

      using (var stream = new FileStream(_clientSecretFile, FileMode.Open, FileAccess.Read))
      {
         credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
             GoogleClientSecrets.FromStream(stream).Secrets,
             new[] { DriveService.Scope.Drive },
             "user",
             CancellationToken.None,
             new FileDataStore(_tokenStorePath, true),
             new ManualCodeReceiver(_owner)
         );
      }

      _service = new DriveService(new BaseClientService.Initializer()
      {
         HttpClientInitializer = credential,
         ApplicationName = "SPM Google Drive"
      });

      return _service;
   }

   public void DisconnectService()
   {
      _service = null;
   }


   // -- ManualCodeReceiver (logowanie bez localhost)
   public class ManualCodeReceiver : ICodeReceiver
   {
      private Form owner;

      public ManualCodeReceiver(Form ownerWindow)
      {
         owner = ownerWindow;
      }

      public string RedirectUri => "urn:ietf:wg:oauth:2.0:oob";

      public Task<AuthorizationCodeResponseUrl> ReceiveCodeAsync(
          AuthorizationCodeRequestUrl url,
          CancellationToken cancellationToken)
      {
         string code = string.Empty;

         // 1. Otwórz przeglądarkę Google
         System.Diagnostics.Process.Start(url.Build().ToString());

         // 2. Pokaż okno instrukcji z właścicielem
         using (Form instructionForm = new Form()
         {
            Text = "Instrukcja",
            Size = new Size(400, 150),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false,
            MinimizeBox = false
         })
         {
            Label lbl = new Label()
            {
               Text = "Otworzyła się przeglądarka Google.\nSkopiuj kod autoryzacyjny i kliknij OK, aby kontynuować.",
               Dock = DockStyle.Fill,
               TextAlign = ContentAlignment.MiddleCenter
            };
            instructionForm.Controls.Add(lbl);

            Button btnOk = new Button()
            {
               Text = "OK",
               Dock = DockStyle.Bottom,
               Height = 35
            };
            btnOk.Click += (s, e) => instructionForm.Close();
            instructionForm.Controls.Add(btnOk);

            instructionForm.ShowDialog(owner); // <--- ustawiamy właściciela
         }

         // 3. Pokaż PasswordDialog z właścicielem
         using (var pd = new PasswordDialog("Wprowadź kod autoryzacyjny Google:"))
         {
            if (pd.ShowDialog(owner) == DialogResult.OK) // <--- ustawiamy właściciela
            {
               code = pd.Password;
            }
         }

         return Task.FromResult(new AuthorizationCodeResponseUrl { Code = code });
      }
   }


   // -- Utwórz foldery wg ścieżki /folder1/folder2/
   public async Task<string> GetOrCreateFolderPathAsync(string path)
   {
      var service = await GetServiceAsync();

      string[] parts = path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
      string parentId = "root";

      foreach (var folderName in parts)
      {
         var list = service.Files.List();
         list.Q =
             $"mimeType='application/vnd.google-apps.folder' " +
             $"and name='{folderName}' " +
             $"and '{parentId}' in parents and trashed=false";
         list.Fields = "files(id,name)";
         var result = await list.ExecuteAsync();

         if (result.Files.Any())
         {
            parentId = result.Files[0].Id;
         }
         else
         {
            var folder = new Google.Apis.Drive.v3.Data.File()
            {
               Name = folderName,
               MimeType = "application/vnd.google-apps.folder",
               Parents = new List<string> { parentId }
            };

            var create = service.Files.Create(folder);
            create.Fields = "id";
            parentId = (await create.ExecuteAsync()).Id;
         }
      }

      return parentId;
   }

   // -- Wysyłanie pliku + nadpisywanie jeśli istnieje
   public async Task UploadFileToPathAsync(string folderPath, string localFilePath)
   {
      var service = await GetServiceAsync();

      string folderId = await GetOrCreateFolderPathAsync(folderPath);
      string fileName = Path.GetFileName(localFilePath);

      // Sprawdzenie istnienia pliku
      var list = service.Files.List();
      list.Q = $"name='{fileName}' and '{folderId}' in parents and trashed=false";
      list.Fields = "files(id)";
      var existing = await list.ExecuteAsync();

      string fileId = existing.Files.FirstOrDefault()?.Id;

      using (var stream = new FileStream(localFilePath, FileMode.Open))
      {
         if (fileId == null)
         {
            // Nowy plik
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
               Name = fileName,
               Parents = new List<string> { folderId }
            };

            var request = service.Files.Create(fileMetadata, stream, "application/octet-stream");
            request.Fields = "id";
            await request.UploadAsync();
         }
         else
         {
            // Nadpisanie
            var fileMetadata = new Google.Apis.Drive.v3.Data.File() { Name = fileName };

            var request = service.Files.Update(fileMetadata, fileId, stream, "application/octet-stream");
            request.Fields = "id";
            await request.UploadAsync();
         }
      }
   }

   // -- Pobranie ID folderu bez tworzenia
   public async Task<string> GetFolderIdByPathAsync(string path)
   {
      var service = await GetServiceAsync();

      string[] parts = path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
      string parentId = "root";

      foreach (var folderName in parts)
      {
         var list = service.Files.List();
         list.Q =
             $"mimeType='application/vnd.google-apps.folder' " +
             $"and name='{folderName}' " +
             $"and '{parentId}' in parents and trashed=false";
         list.Fields = "files(id,name)";

         var res = await list.ExecuteAsync();

         if (!res.Files.Any())
            return null;

         parentId = res.Files[0].Id;
      }

      return parentId;
   }

   // -- Pobranie pliku
   public async Task DownloadFileFromPathAsync(string folderPath, string fileName, string savePath)
   {
      var service = await GetServiceAsync();

      string folderId = await GetFolderIdByPathAsync(folderPath);
      if (folderId == null)
         throw new Exception("Folder nie istnieje: " + folderPath);

      var list = service.Files.List();
      list.Q = $"name='{fileName}' and '{folderId}' in parents and trashed=false";
      list.Fields = "files(id,name)";
      var files = await list.ExecuteAsync();

      var file = files.Files.FirstOrDefault();
      if (file == null)
         throw new Exception("Nie znaleziono pliku: " + fileName);

      var request = service.Files.Get(file.Id);

      using (var stream = new FileStream(savePath, FileMode.Create))
      {
         await request.DownloadAsync(stream);
      }
   }

   // -- Lista plików w folderze
   public async Task<IList<Google.Apis.Drive.v3.Data.File>> GetFilesInFolderAsync(string folderPath)
   {
      var service = await GetServiceAsync();

      string folderId = await GetFolderIdByPathAsync(folderPath);
      if (folderId == null)
         return new List<Google.Apis.Drive.v3.Data.File>();

      var list = service.Files.List();
      list.Q = $"'{folderId}' in parents and trashed=false";
      list.Fields = "files(id,name,mimeType)";

      return (await list.ExecuteAsync()).Files;
   }

   // -- Usuwanie pliku jeśli istnieje
   public async Task DeleteFileIfExistsAsync(string folderPath, string fileName)
   {
      var service = await GetServiceAsync();

      string folderId = await GetFolderIdByPathAsync(folderPath);
      if (folderId == null)
         return;

      var list = service.Files.List();
      list.Q = $"name='{fileName}' and '{folderId}' in parents and trashed=false";
      list.Fields = "files(id)";
      var files = await list.ExecuteAsync();

      foreach (var file in files.Files)
      {
         await service.Files.Delete(file.Id).ExecuteAsync();
      }
   }
}
