namespace Awr.Worker.Configuration
{
    public static class WorkerConstants
    {
        // Paths
        public const string SourceLocation = @"C:\AwrTest\SourceTemplates";
        public const string TempLocation = @"C:\AwrTest\Working";
        public const string FinalLocation = @"C:\AwrTest\FinalIssuedDocs"; // Where QA generated docs live

        // Security
        public const string EncryptionPassword = "QA";
        public const string RestrictEditPassword = "test123";

        // Resilience
        public const int MaxRetries = 3;

        // IPC Codes
        public const int SuccessExitCode = 0;
        public const int FailureExitCode = 1;
        public const string DateTimeFormat = "dd-MM-yyyy HH:mm:ss";

        // --- NEW: Modes ---
        public const string ModeGenerate = "GENERATE";
        public const string ModePrint = "PRINT";
    }
}