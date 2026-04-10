using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace AudioRoute;

internal sealed class StaThreadDispatcher : IDisposable
{
    private readonly BlockingCollection<IWorkItem> workItems = new();
    private readonly Thread thread;
    private volatile bool disposed;
    private int managedThreadId;

    public StaThreadDispatcher(string threadName)
    {
        thread = new Thread(Run)
        {
            IsBackground = true,
            Name = threadName
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task<T> InvokeAsync<T>(Func<T> work)
    {
        if (disposed)
            throw new ObjectDisposedException(nameof(StaThreadDispatcher));

        if (Environment.CurrentManagedThreadId == managedThreadId)
        {
            try
            {
                return Task.FromResult(work());
            }
            catch (Exception ex)
            {
                return Task.FromException<T>(ex);
            }
        }

        var workItem = new WorkItem<T>(work);
        workItems.Add(workItem);
        return workItem.Task;
    }

    public T Invoke<T>(Func<T> work)
    {
        return InvokeAsync(work).GetAwaiter().GetResult();
    }

    public void Invoke(Action work)
    {
        Invoke(() =>
        {
            work();
            return true;
        });
    }

    public void Dispose()
    {
        if (disposed)
            return;

        disposed = true;
        workItems.CompleteAdding();

        if (Environment.CurrentManagedThreadId != managedThreadId && thread.IsAlive)
            thread.Join(TimeSpan.FromSeconds(2));
    }

    private void Run()
    {
        managedThreadId = Environment.CurrentManagedThreadId;

        foreach (var workItem in workItems.GetConsumingEnumerable())
            workItem.Execute();
    }

    private interface IWorkItem
    {
        void Execute();
    }

    private sealed class WorkItem<T> : IWorkItem
    {
        private readonly Func<T> work;
        private readonly TaskCompletionSource<T> completionSource = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public WorkItem(Func<T> work)
        {
            this.work = work;
        }

        public Task<T> Task => completionSource.Task;

        public void Execute()
        {
            try
            {
                completionSource.TrySetResult(work());
            }
            catch (Exception ex)
            {
                completionSource.TrySetException(ex);
            }
        }
    }
}
