namespace VCardFileCleaner{
public class VCardRecordComparer : IEqualityComparer<VCardRecord>
{
    public bool Equals(VCardRecord? x, VCardRecord? y)
    {
        return x?.Tel == y?.Tel && x?.Tel2 == y?.Tel2 && x?.Tel3 == y?.Tel3;
    }

    public int GetHashCode(VCardRecord obj)
    {
        return obj?.Tel?.GetHashCode() ^ obj?.Tel2?.GetHashCode() ^ obj?.Tel3?.GetHashCode() ?? 0;
    }
}}