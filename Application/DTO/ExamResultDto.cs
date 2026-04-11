namespace Application.DTO;

public record ExamResultDto(
    string FirstName,
    string LastName,
    List<ExamDto> Exam,
    string Club);

public record ExamDto(
    string ExamName, 
    float Grade, 
    string Tendency, 
    string Examiner, 
    string Comment, 
    bool Marked);