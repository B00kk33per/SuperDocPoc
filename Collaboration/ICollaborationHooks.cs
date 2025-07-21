namespace SuperDocPoc.Collaboration;

public interface ICollaborationHooks
{
    Task<bool> AuthenticateAsync(HttpContext context);
    Task<byte[]?> LoadDocumentAsync(string documentId);
    Task SaveDocumentAsync(string documentId, byte[] state);
    Task OnChangeAsync(string documentId, byte[] update);
}
