namespace Awr.Core.Enums
{
    // Defines the possible states for a single AWR line item
    public enum AwrItemStatus
    {
        Draft = 0,
        PendingIssuance = 1, // Ready for QA to issue
        Issued = 2,          // QA has stamped and issued
        Received = 3,        // Requester/QC has received, waiting for return (Step 3 complete)
        PendingRetrieval = 4, // Requester/QC has returned, waiting for QA check (Step 4 complete)
        Complete = 5,        // QA has retrieved and closed the item (Step 5 complete)
        RejectedByQa = 6     // Rejected at the QA stage
    }
}