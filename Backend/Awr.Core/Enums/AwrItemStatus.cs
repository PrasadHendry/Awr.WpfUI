namespace Awr.Core.Enums
{
    public enum AwrItemStatus
    {
        Draft = 0,
        PendingIssuance = 1,
        Issued = 2,
        Received = 3, // InUse / Printed
        Voided = 4,   // Replaces 'PendingRetrieval'
        Complete = 5,
        RejectedByQa = 6
    }
}