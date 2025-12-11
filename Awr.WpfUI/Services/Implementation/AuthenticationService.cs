using Awr.Core.DTOs;
using Awr.WpfUI.Services.Interfaces;
using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace Awr.WpfUI.Services.Implementation
{
    public class AuthenticationService : IAuthenticationService
    {
        private readonly string _connectionString;

        // Hardcoded Fallback (Used only if DB row is missing entirely)
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
                const string sql = @"
                    SELECT RoleID, RoleName 
                    FROM dbo.User_Roles 
                    ORDER BY CASE RoleName WHEN 'Requester' THEN 1 WHEN 'QA' THEN 2 WHEN 'Admin' THEN 3 ELSE 4 END;";

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

                // 1. Check Database FIRST
                const string sql = "SELECT PasswordHash FROM dbo.User_Roles WHERE RoleName = @RoleName";

                using (var connection = new SqlConnection(_connectionString))
                {
                    string dbHash = connection.QuerySingleOrDefault<string>(sql, new { RoleName = roleName });

                    // If user exists in DB, check DB password
                    if (dbHash != null)
                    {
                        if (dbHash == password) return (true, roleName);
                        return (false, string.Empty); // Wrong password
                    }
                }

                // 2. Fallback: If user NOT in DB (shouldn't happen in prod), check hardcoded
                if (AbsoluteCredentials.TryGetValue(roleName, out string hardHash))
                {
                    if (hardHash == password) return (true, roleName);
                }

                return (false, string.Empty);
            });
        }

        public async Task UpdatePasswordAsync(string roleName, string newPassword)
        {
            await Task.Run(() =>
            {
                const string sql = "UPDATE dbo.User_Roles SET PasswordHash = @Pwd WHERE RoleName = @RoleName";
                using (var connection = new SqlConnection(_connectionString))
                {
                    connection.Execute(sql, new { Pwd = newPassword, RoleName = roleName });
                }
            });
        }
    }
}