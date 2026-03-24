using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace AMTool;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    private const string SingleInstanceMutexName = @"Local\AMTool.SingleInstance";
    private const string ActivateExistingInstanceEventName = @"Local\AMTool.ActivateExistingInstance";

    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _activateExistingInstanceEvent;
    private RegisteredWaitHandle? _activateExistingInstanceRegistration;
    private bool _ownsSingleInstanceMutex;

    public App()
    {
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        try
        {
            if (!TryAcquireSingleInstanceGuard())
            {
                NotifyRunningInstance();
                Shutdown();
                return;
            }

            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Show();
        }
        catch (Exception exception)
        {
            HandleUnexpectedException(exception, "Nie udalo sie uruchomic aplikacji.");
            Shutdown(-1);
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        ReleaseSingleInstanceGuard();
        base.OnExit(e);
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        HandleUnexpectedException(e.Exception, "Wystapil nieoczekiwany blad aplikacji.");
        e.Handled = true;
    }

    private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        HandleUnexpectedException(e.ExceptionObject as Exception, "Aplikacja napotkala krytyczny blad.");
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        HandleUnexpectedException(e.Exception, "Wystapil blad w tle aplikacji.");
        e.SetObserved();
    }

    private bool TryAcquireSingleInstanceGuard()
    {
        _singleInstanceMutex = new Mutex(initiallyOwned: true, SingleInstanceMutexName, out bool createdNew);

        if (!createdNew)
        {
            _singleInstanceMutex.Dispose();
            _singleInstanceMutex = null;
            _ownsSingleInstanceMutex = false;
            return false;
        }

        _ownsSingleInstanceMutex = true;
        _activateExistingInstanceEvent = new EventWaitHandle(
            initialState: false,
            EventResetMode.AutoReset,
            ActivateExistingInstanceEventName);

        _activateExistingInstanceRegistration = ThreadPool.RegisterWaitForSingleObject(
            _activateExistingInstanceEvent,
            OnActivateExistingInstanceRequested,
            state: null,
            Timeout.Infinite,
            executeOnlyOnce: false);

        return true;
    }

    private void NotifyRunningInstance()
    {
        for (int attempt = 0; attempt < 20; attempt++)
        {
            try
            {
                using EventWaitHandle activationEvent = EventWaitHandle.OpenExisting(ActivateExistingInstanceEventName);
                activationEvent.Set();
                return;
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                Thread.Sleep(100);
            }
            catch
            {
                return;
            }
        }
    }

    private void OnActivateExistingInstanceRequested(object? state, bool timedOut)
    {
        if (timedOut)
        {
            return;
        }

        _ = Dispatcher.BeginInvoke(new Action(() => _ = TryActivateMainWindowAsync()));
    }

    private async Task TryActivateMainWindowAsync()
    {
        try
        {
            if (MainWindow is AMTool.MainWindow mainWindow)
            {
                await mainWindow.ActivateFromExternalRequestAsync();
            }
        }
        catch (Exception exception)
        {
            TryLogException(exception);
        }
    }

    private static void HandleUnexpectedException(Exception? exception, string message)
    {
        try
        {
            TryLogException(exception);

            string details = exception?.GetBaseException().Message ?? "Brak dodatkowych szczegolow.";
            System.Windows.MessageBox.Show(
                $"{message}{Environment.NewLine}{Environment.NewLine}{details}",
                "AMTool",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        catch
        {
        }
    }

    private static void TryLogException(Exception? exception)
    {
        if (exception is null)
        {
            return;
        }

        try
        {
            string appDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AMTool");

            Directory.CreateDirectory(appDataDirectory);

            string logPath = Path.Combine(appDataDirectory, "exceptions.log");
            string logEntry = string.Join(
                Environment.NewLine,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}]",
                exception.ToString(),
                new string('-', 80),
                string.Empty);

            File.AppendAllText(logPath, logEntry);
        }
        catch
        {
        }
    }

    private void ReleaseSingleInstanceGuard()
    {
        try
        {
            _activateExistingInstanceRegistration?.Unregister(null);
        }
        catch
        {
        }
        finally
        {
            _activateExistingInstanceRegistration = null;
        }

        if (_activateExistingInstanceEvent is not null)
        {
            _activateExistingInstanceEvent.Dispose();
            _activateExistingInstanceEvent = null;
        }

        if (_singleInstanceMutex is null)
        {
            return;
        }

        try
        {
            if (_ownsSingleInstanceMutex)
            {
                _singleInstanceMutex.ReleaseMutex();
            }
        }
        catch
        {
        }
        finally
        {
            _singleInstanceMutex.Dispose();
            _singleInstanceMutex = null;
            _ownsSingleInstanceMutex = false;
        }
    }
}
