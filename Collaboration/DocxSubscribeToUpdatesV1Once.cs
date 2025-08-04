using YDotNet.Document;

namespace SuperDocPoc.Collaboration;

/// <summary>
/// Custom implementation of SubscribeToUpdatesV1Once for SuperDocPoc
/// Based on YDotNet.Server's internal SubscribeToUpdatesV1Once
/// </summary>
internal sealed class DocxSubscribeToUpdatesV1Once : IDisposable
{
    private readonly IDisposable unsubscribe;

    public DocxSubscribeToUpdatesV1Once(Doc doc)
    {
        unsubscribe = doc.ObserveUpdatesV1(@event =>
        {
            Update = @event.Update;
        });
    }

    public byte[]? Update { get; private set; }

    public void Dispose()
    {
        unsubscribe.Dispose();
    }
}
