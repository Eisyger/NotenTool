namespace Domain.Entity;

public class ExamResult(Exam exam, Student student, float grade)
{
    public Exam Exam { get; init; } = exam;
    public Student Student { get; init; } = student;
    public float Grade { get; set; } = grade;
    
    public string Tendency { get; set; } = string.Empty;
    
    public string Comment { get; set; } = string.Empty;
    
    public bool IsMarked { get; set; } = false;
    
    public string Examiner { get; set; } = string.Empty;

    public static List<ExamResult> CreateDefault(List<Student> students, List<Exam> exams)
    {
        return (from e in exams from s in students select new ExamResult(e, s, 0.0f)).ToList();
    }
}