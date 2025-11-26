using System.Collections.Generic;
using System.Threading.Tasks;
using Awr.Core.DTOs;

namespace Awr.WpfUI.Services.Interfaces
{
    public interface IAuthenticationService
    {
        Task<(bool IsSuccess, string Role)> ValidateUserAsync(string username, string password);
        Task<List<UserRoleDto>> GetUserRolesAsync();
    }
}