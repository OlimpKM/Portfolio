using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Aplikacja
{
   public partial class ucWebView2 : UserControl
   {
      private CancellationTokenSource _navigationTimeoutCts;
      private bool _FlagBrowserReady = false;
      private bool _FlagLoadingPage = false;

      public event EventHandler<MessageEventArgs> WebMessage;
      public event EventHandler<MessageEventArgs> WebError;
      public event EventHandler<MessageEventArgs> MessageReceived;
      public event Action<object> InitializationCompleted;
      public event EventHandler<BeforeLoadEventArgs> BeforeLoadPage;
      public event EventHandler<AfterLoadEventArgs> AfterLoadPage;
      public event EventHandler<SourceChangedEventArgs> SourceChanged;

      public ucWebView2()
      {
         InitializeComponent();
         webView.CoreWebView2InitializationCompleted += WebView_CoreWebView2InitializationCompleted;

         // Uruchomienie asynchronicznej inicjalizacji
         _ = InitializeWebViewAsync();
      }

      private async Task InitializeWebViewAsync()
      {
         var timeoutTask = Task.Delay(3000);
         var readyTask = EnsureWebViewReadyAsync(webView);

         // Oczekiwanie na pierwsze zakończone zadanie
         await Task.WhenAny(readyTask, timeoutTask);
      }

      private async Task EnsureWebViewReadyAsync(WebView2 webView)
      {
         try
         {
            if (webView?.CoreWebView2 == null) await webView.EnsureCoreWebView2Async();
         }
         catch (Exception ex)
         {
            OnError(new MessageEventArgs($"Wystąpił błąd: {ex.Message}"));
         }
      }

      protected override void Dispose(bool disposing)
      {
         if (disposing && (components != null))
         {
            _navigationTimeoutCts?.Cancel();
            _navigationTimeoutCts?.Dispose();
            _navigationTimeoutCts = null;

            if (webView != null)
            {
               webView.CoreWebView2InitializationCompleted -= WebView_CoreWebView2InitializationCompleted;
               if (webView.CoreWebView2 != null)
               {
                  webView.CoreWebView2.WebMessageReceived -= WebView_WebMessageReceived;
                  webView.CoreWebView2.SourceChanged -= WebView_SourceChanged;
                  webView.CoreWebView2.NavigationStarting -= WebView_NavigationStarting;
                  webView.CoreWebView2.NavigationCompleted -= WebView_NavigationCompleted;
                  webView.CoreWebView2.Stop();
               }
               webView.Dispose();
            }
            components.Dispose();
         }
         base.Dispose(disposing);
      }

      protected virtual void OnMessage(MessageEventArgs e)
      {
         WebMessage?.Invoke(webView, e);
      }
      protected virtual void OnError(MessageEventArgs e)
      {
         WebError?.Invoke(webView, e);
      }
      protected virtual void OnInitializationCompleted()
      {
         InitializationCompleted?.Invoke(webView);
      }

      protected virtual void OnMessageReceived(MessageEventArgs e)
      {
         MessageReceived?.Invoke(webView,e);
      }

      protected virtual void OnBeforeLoadPage(BeforeLoadEventArgs e)
      {
         BeforeLoadPage?.Invoke(webView, e);
      }

      protected virtual void OnAfterLoadPage(AfterLoadEventArgs e)
      {
         AfterLoadPage?.Invoke(webView, e);
      }

      protected virtual void OnSourceChanged(SourceChangedEventArgs e)
      {
         SourceChanged?.Invoke(webView, e);
      }

      public bool IsBrowserReady { get { return _FlagBrowserReady; } }
      public bool IsLoadingPage { get { return _FlagLoadingPage; } }

      public WebView2 WebBrowser { get { return webView; } }

      public void LoadPage(string address)
      {
         if (_FlagBrowserReady && !_FlagLoadingPage)
            StartNavigationWithTimeout(address);
         else
            OnError(new MessageEventArgs($"Przeglądarka nie gotowa"));
      }

      public void EmptyPage()
      {
         if (_FlagBrowserReady && !_FlagLoadingPage)
            StartNavigationWithTimeout("about:blank");
         else
            OnError(new MessageEventArgs($"Przeglądarka nie gotowa"));
      }

      public void RealodPage()
      {
         if (_FlagBrowserReady && !_FlagLoadingPage)
            webView.CoreWebView2.Reload();
         else
            OnError(new MessageEventArgs($"Przeglądarka nie gotowa"));
      }

      public void CancelLoad()
      {
         _navigationTimeoutCts?.Cancel();
      }

      public void GoBack()
      {
         if (_FlagBrowserReady && !_FlagLoadingPage)
         webView.CoreWebView2.GoBack();
      }


      private void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
      {
         if (e.IsSuccess)
         {
            _FlagBrowserReady = true;
            webView.WebMessageReceived += WebView_WebMessageReceived;
            webView.CoreWebView2.SourceChanged += WebView_SourceChanged;
            webView.CoreWebView2.NavigationStarting += WebView_NavigationStarting;
            webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;

            OnInitializationCompleted();
         }
         else
            OnError(new MessageEventArgs($"Niepowodzenie inicjalizacji WebView2: {e.InitializationException.Message}"));
      }

      private void WebView_NavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs e)
      {
         _FlagLoadingPage = true;
         OnBeforeLoadPage(new BeforeLoadEventArgs(webView.Source.ToString()));
      }

      private void WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
      {
         _FlagLoadingPage = false;
         OnAfterLoadPage(new AfterLoadEventArgs(webView.Source.ToString(), true));
      }

      private void WebView_SourceChanged(object sender, CoreWebView2SourceChangedEventArgs e)
      {
         OnSourceChanged(new SourceChangedEventArgs(webView.Source.ToString()));
      }

      private void WebView_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
      {
         OnMessageReceived(new MessageEventArgs(e.TryGetWebMessageAsString()));
      }

      private async void StartNavigationWithTimeout(string url, int timeoutMilliseconds = 5000)
      {
         try
         {
            _navigationTimeoutCts = new CancellationTokenSource();
            var navigationTask = NavigateAsync(url, _navigationTimeoutCts.Token);

            if (await Task.WhenAny(navigationTask, Task.Delay(timeoutMilliseconds)) == navigationTask)
            {
               // Sprawdzanie statusu zadania
               if (navigationTask.Status == TaskStatus.RanToCompletion)
                  OnMessage(new MessageEventArgs("Strona została załadowana pomyślnie."));
               else if (navigationTask.Status == TaskStatus.Faulted)
                  OnMessage(new MessageEventArgs("Wystąpił błąd podczas ładowania strony."));
               else if (navigationTask.Status == TaskStatus.Canceled)
                  OnMessage(new MessageEventArgs("Operacja ładowania strony została anulowana."));
            }
            else
            {
               _navigationTimeoutCts.Cancel();
               webView.CoreWebView2?.Stop();

               OnMessage(new MessageEventArgs("Przekroczono czas ładowania. Ładowanie innej strony."));
               webView.CoreWebView2.Navigate("about:blank");
            }
         }
         catch (Exception ex)
         {
            OnError(new MessageEventArgs($"Wystąpił błąd: {ex.Message}"));
         }
         finally
         {
            _navigationTimeoutCts?.Dispose();
            _navigationTimeoutCts = null;
         }
      }

      private async Task NavigateAsync(string url, CancellationToken token)
      {
         try
         {
            webView.CoreWebView2.Navigate(url);
            var tcs = new TaskCompletionSource<object>();

            EventHandler<CoreWebView2NavigationCompletedEventArgs> handler = null;
            handler = (s, e) =>
            {
               webView.CoreWebView2.NavigationCompleted -= handler;
               if (!token.IsCancellationRequested)
               {
                  if (e.IsSuccess)
                     tcs.SetResult(null);
                  else
                     tcs.SetException(new Exception($"Błąd podczas ładowania: {e.WebErrorStatus}"));
               }
            };
            webView.CoreWebView2.NavigationCompleted += handler;
            await tcs.Task;
         }
         catch (OperationCanceledException ex)
         {
            OnError(new MessageEventArgs($"Wystąpił błąd: {ex.Message}"));
         }
      }
   }

   public class MessageEventArgs : EventArgs
   {
      public string Message { get; }

      public MessageEventArgs(string message)
      {
         Message = message;
      }
   }

   public class BeforeLoadEventArgs : EventArgs
   {
      public string Address { get; }

      public BeforeLoadEventArgs(string message)
      {
         Address = message;
      }
   }

   public class AfterLoadEventArgs : EventArgs
   {
      public string Address { get; }
      public bool Completed { get; }

      public AfterLoadEventArgs(string message, bool completed)
      {
         Address = message;
         Completed = completed;
      }
   }

   public class SourceChangedEventArgs : EventArgs
   {
      public string Address { get; }

      public SourceChangedEventArgs(string message)
      {
         Address = message;
      }
   }

}
