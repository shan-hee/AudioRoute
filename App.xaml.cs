using Microsoft.UI.Xaml;

namespace AudioRoute;

public partial class App : Application
{
    private MainWindow? mainWindow;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        var showOnLaunch = !StartupManager.IsStartupLaunch(args.Arguments);
        mainWindow = new MainWindow(showOnLaunch);
        mainWindow.Activate();
    }
}
