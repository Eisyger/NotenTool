namespace Domain.Entity;

public class Course
{
    public List<Student> Students { get; set; } = [];
    public List<Exam> Exams { get; set; } = [];

    public Teacher Teacher { get; set; } = new("blank");

    public CourseMetaData MetaData { get; set; } = new() {FileName = "unknown", LastModifiedDate = DateTime.MinValue, UploadDate = DateTime.MinValue};
}
