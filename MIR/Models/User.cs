namespace MIR.Models
{
    public enum UserRole
    {
        Admin,
        Staff
    }

    public class User
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Username { get; set; } = string.Empty;
        public string PasswordHash { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public UserRole Role { get; set; } = UserRole.Staff;
        public bool IsActive { get; set; } = true;
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public User() { }

        public User(string username, string passwordHash, string fullName, UserRole role)
        {
            Username = username;
            PasswordHash = passwordHash;
            FullName = fullName;
            Role = role;
        }
    }
}
