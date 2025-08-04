using System.Reactive;
using System.Reactive.Concurrency;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Reactive.Subjects;
using System.Reactive.Threading.Tasks;

namespace SuperDocPoc.Collaboration;

/// <summary>
/// Custom DelayedWriter for SuperDocPoc based on YDotNet.Server's internal DelayedWriter
/// </summary>
internal sealed class DocxDelayedWriter
{
    private readonly Subject<bool> writes = new();
    private readonly Task completion;

    public DocxDelayedWriter(TimeSpan delayTime, TimeSpan delayMax, Func<Task> action)
    {
        completion = ThrottleMax(writes, delayTime, delayMax)
            .StartWith(false)
            .Select(write =>
                write
                ? Observable.FromAsync(async () =>
                {
                    await Task.Yield();
                    await action().ConfigureAwait(false);
                })
                : Observable.Return(Unit.Default))
            .Concat()
            .ToTask();
    }

    public Task FlushAsync()
    {
        writes.OnCompleted();
        return completion;
    }

    public void Ping()
    {
        writes.OnNext(value: true);
    }

    private static IObservable<T> ThrottleMax<T>(IObservable<T> source, TimeSpan dueTime, TimeSpan maxTime)
    {
        var scheduler = Scheduler.Default;

        return Observable.Create<T>(o =>
        {
            var lastValue = default(T);
            var lastWrite = scheduler.Now;
            var dueTimeDisposable = new SerialDisposable();

            void Next(T value)
            {
                dueTimeDisposable.Disposable = Disposable.Empty;

                o.OnNext(value);
                lastValue = default;
                lastWrite = scheduler.Now;
            }

            void NextLast()
            {
                Next(lastValue);
            }

            return source.Subscribe(
                x =>
                {
                    if (scheduler.Now - lastWrite > maxTime && dueTimeDisposable.Disposable != Disposable.Empty)
                    {
                        Next(x);
                    }
                    else
                    {
                        lastValue = x;
                        dueTimeDisposable.Disposable = scheduler.Schedule(dueTime, NextLast);
                    }
                },
                o.OnError,
                o.OnCompleted);
        });
    }
}
