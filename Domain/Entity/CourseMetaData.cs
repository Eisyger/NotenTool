namespace Domain.Entity;

public class CourseMetaData
{
    public string FileName { get; set; } = string.Empty;
    public DateTime LastModifiedDate { get; set; }
    public DateTime UploadDate { get; set; }
}