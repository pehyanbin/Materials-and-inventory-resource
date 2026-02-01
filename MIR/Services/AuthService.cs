using MIR.Models;

namespace MIR.Services
{
    public interface IAuthService
    {
        User? CurrentUser { get; }
        bool Login(string username, string password);
        void Logout();
        bool ChangePassword(string oldPassword, string newPassword);
        void SetDefaultUser();
    }

    public class AuthService : IAuthService
    {
        private readonly IExcelService _excelService;
        public User? CurrentUser { get; private set; }

        public AuthService(IExcelService excelService)
        {
            _excelService = excelService;
        }

        public bool Login(string username, string password)
        {
            var user = _excelService.GetUserByUsername(username);
            if (user != null && user.IsActive && BCrypt.Net.BCrypt.Verify(password, user.PasswordHash))
            {
                CurrentUser = user;
                return true;
            }
            return false;
        }

        public void Logout()
        {
            CurrentUser = null;
        }

        public bool ChangePassword(string oldPassword, string newPassword)
        {
            if (CurrentUser == null) return false;

            if (BCrypt.Net.BCrypt.Verify(oldPassword, CurrentUser.PasswordHash))
            {
                CurrentUser.PasswordHash = BCrypt.Net.BCrypt.HashPassword(newPassword);
                _excelService.UpdateUser(CurrentUser);
                return true;
            }
            return false;
        }

        public void SetDefaultUser()
        {
            CurrentUser = _excelService.GetUserByUsername("admin");
        }
    }
}
