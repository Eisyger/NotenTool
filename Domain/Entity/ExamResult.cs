namespace Domain.Entity;

public class ExamResult(Exam exam, Student student, float grade)
{
    public Exam Exam { get; init; } = exam;
    public Student Student { get; init; } = student;
    public float Grade { get; set; } = grade;
    
    public string Tendency { get; set; } = string.Empty;
    
    public string Comment { get; set; } = string.Empty;
}