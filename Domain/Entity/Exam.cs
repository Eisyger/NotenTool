namespace Domain.Entity;

public class Exam
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;

    public ExamResult Examine(Student student, float grade)
    {
        return new ExamResult(this, student, grade);
    }
}