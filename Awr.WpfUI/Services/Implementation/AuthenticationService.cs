using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Awr.Core.DTOs;
using Awr.WpfUI.Services.Interfaces;
using Dapper;

namespace Awr.WpfUI.Services.Implementation
{
    public class AuthenticationService : IAuthenticationService
    {
        private readonly string _connectionString;

        private static readonly Dictionary<string, string> AbsoluteCredentials = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Requester", "test1" },
            { "QA", "test2" },
            { "Admin", "adminpass" }
        };

        public AuthenticationService()
        {
            _connectionString = ConfigurationManager.ConnectionStrings["AwrDbConnection"]?.ConnectionString;
        }

        public async Task<List<UserRoleDto>> GetUserRolesAsync()
        {
            return await Task.Run(() =>
            {
                // FIX: Custom Sort Order (Requester -> QA -> Admin)
                const string sql = @"
                    SELECT RoleID, RoleName 
                    FROM dbo.User_Roles 
                    ORDER BY 
                        CASE RoleName 
                            WHEN 'Requester' THEN 1 
                            WHEN 'QA' THEN 2 
                            WHEN 'Admin' THEN 3 
                            ELSE 4 
                        END;";

                using (IDbConnection connection = new SqlConnection(_connectionString))
                {
                    return connection.Query<UserRoleDto>(sql).ToList();
                }
            });
        }

        public async Task<(bool IsSuccess, string Role)> ValidateUserAsync(string username, string password)
        {
            return await Task.Run(() =>
            {
                string roleName = username.Trim();

                if (!AbsoluteCredentials.TryGetValue(roleName, out string expectedHash))
                    return (false, string.Empty);

                if (password != expectedHash)
                    return (false, string.Empty);

                const string sql = @"
                    SELECT RoleName 
                    FROM dbo.User_Roles 
                    WHERE RoleName = @RoleName AND PasswordHash = @ExpectedHash;";

                using (IDbConnection connection = new SqlConnection(_connectionString))
                {
                    var parameters = new { RoleName = roleName, ExpectedHash = expectedHash };
                    string userRole = connection.QuerySingleOrDefault<string>(sql, parameters);

                    if (!string.IsNullOrEmpty(userRole))
                        return (true, userRole);
                }

                return (false, string.Empty);
            });
        }
    }
}