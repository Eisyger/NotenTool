namespace NotenTool.DTO;

public class FilePickerResult
{
    public string FileName { get; set; } = string.Empty;
    public long LastModified { get; set; }
    public long Size { get; set; }
    public string Type { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
    
    public DateTime LastModifiedDate => 
        DateTimeOffset.FromUnixTimeMilliseconds(LastModified).DateTime;

    public DateTime UploadDate { get; set; }
    
    public byte[] GetBytes() => Convert.FromBase64String(Content);
}