namespace Domain.Entity;

public class Student
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public string Club { get; set; } = string.Empty;
    public DateTime DateOfBirth { get; set; } = DateTime.Now;
}