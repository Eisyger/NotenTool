namespace Domain.Entity;

public class CourseMetaData
{
    public string FileName
    {
        get => _filename;
        init => _filename = Path.GetFileNameWithoutExtension(value);
    }

    private string _filename = string.Empty;

    public DateTime LastModifiedDate { get; set; }
    public DateTime UploadDate { get; set; }
}