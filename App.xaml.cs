using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace AMTool;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    public App()
    {
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
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
}
