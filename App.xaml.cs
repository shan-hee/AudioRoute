using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;

namespace AudioRoute;

public partial class App : Application
{
    private const string SingleInstanceMutexName = @"Local\AudioRoute.Singleton";
    private const string SingleInstanceActivationEventName = @"Local\AudioRoute.Singleton.Activate";

    private readonly EventWaitHandle activationEvent;
    private readonly CancellationTokenSource activationMonitorCancellation = new();
    private readonly bool isPrimaryInstance;
    private readonly Mutex singleInstanceMutex;
    private MainWindow? mainWindow;
    private int pendingExternalActivationCount;
    private int singleInstanceResourcesDisposed;

    public App()
    {
        singleInstanceMutex = new Mutex(initiallyOwned: true, SingleInstanceMutexName, out var createdNew);
        isPrimaryInstance = createdNew;
        activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, SingleInstanceActivationEventName);
        InitializeComponent();

        if (isPrimaryInstance)
            _ = Task.Run(MonitorExternalActivationAsync);
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        if (!isPrimaryInstance)
        {
            if (!StartupManager.IsStartupLaunch(args.Arguments))
                SignalPrimaryInstanceActivation();

            DisposeSingleInstanceResources();
            Exit();
            return;
        }

        var showOnLaunch = !StartupManager.IsStartupLaunch(args.Arguments);
        mainWindow = new MainWindow();
        mainWindow.Closed += (_, _) => DisposeSingleInstanceResources();

        if (showOnLaunch || Interlocked.Exchange(ref pendingExternalActivationCount, 0) > 0)
            _ = mainWindow.ShowOrBringToFrontAsync();
    }

    private Task MonitorExternalActivationAsync()
    {
        var waitHandles = new WaitHandle[]
        {
            activationEvent,
            activationMonitorCancellation.Token.WaitHandle
        };

        while (true)
        {
            var signaledIndex = WaitHandle.WaitAny(waitHandles);
            if (signaledIndex != 0)
                return Task.CompletedTask;

            if (mainWindow?.DispatcherQueue is DispatcherQueue dispatcherQueue)
            {
                _ = dispatcherQueue.TryEnqueue(() => _ = mainWindow.ShowOrBringToFrontAsync());
                continue;
            }

            Interlocked.Increment(ref pendingExternalActivationCount);
        }
    }

    private void SignalPrimaryInstanceActivation()
    {
        try
        {
            _ = activationEvent.Set();
        }
        catch
        {
        }
    }

    private void DisposeSingleInstanceResources()
    {
        if (Interlocked.Exchange(ref singleInstanceResourcesDisposed, 1) != 0)
            return;

        activationMonitorCancellation.Cancel();

        if (isPrimaryInstance)
            singleInstanceMutex.ReleaseMutex();

        singleInstanceMutex.Dispose();
    }
}
