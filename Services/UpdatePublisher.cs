using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace NetworkAdapterSwitcher.Services;

internal static class UpdatePublisher
{
    private static readonly Guid Authenticode = new("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");

    public static bool IsTrusted(string path) => TrustedKey(path) is not null;

    public static bool MatchesInstalledPublisher(string installed, string update)
    {
        var expected = TrustedKey(installed);
        var actual = TrustedKey(update);
        return expected is not null && actual is not null &&
            CryptographicOperations.FixedTimeEquals(expected, actual);
    }

    private static byte[]? TrustedKey(string path)
    {
        var fileInfo = new FileInfoData { Size = (uint)Marshal.SizeOf<FileInfoData>(), Path = path };
        var file = Marshal.AllocHGlobal(Marshal.SizeOf<FileInfoData>());
        Marshal.StructureToPtr(fileInfo, file, false);
        var action = Authenticode;
        var data = new TrustData
        {
            Size = (uint)Marshal.SizeOf<TrustData>(),
            UiChoice = 2, // WTD_UI_NONE
            RevocationChecks = 1, // WTD_REVOKE_WHOLECHAIN
            UnionChoice = 1, // WTD_CHOICE_FILE
            File = file,
            StateAction = 1, // WTD_STATEACTION_VERIFY
            ProviderFlags = 0x80 // WTD_REVOCATION_CHECK_CHAIN_EXCLUDE_ROOT
        };
        try
        {
            if (WinVerifyTrust(new IntPtr(-1), ref action, ref data) != 0)
                return null;
            var provider = WTHelperProvDataFromStateData(data.StateData);
            if (provider == IntPtr.Zero)
                return null;
            var signer = WTHelperGetProvSignerFromChain(provider, 0, false, 0);
            if (signer == IntPtr.Zero)
                return null;
            var entry = WTHelperGetProvCertFromChain(signer, 0);
            if (entry == IntPtr.Zero)
                return null;
            var certificate = Marshal.PtrToStructure<TrustCertificate>(entry).Context;
            if (certificate == IntPtr.Zero)
                return null;

            var context = Marshal.PtrToStructure<CertificateContext>(certificate);
            if (context.Encoded == IntPtr.Zero || context.EncodedLength is 0 or > 65536)
                return null;
            var bytes = new byte[context.EncodedLength];
            Marshal.Copy(context.Encoded, bytes, 0, bytes.Length);
            using var trustedSigner = X509CertificateLoader.LoadCertificate(bytes);
            return trustedSigner.PublicKey.ExportSubjectPublicKeyInfo();
        }
        catch (CryptographicException) { return null; }
        finally
        {
            if (data.StateData != IntPtr.Zero)
            {
                data.StateAction = 2; // WTD_STATEACTION_CLOSE
                _ = WinVerifyTrust(new IntPtr(-1), ref action, ref data);
            }
            Marshal.DestroyStructure<FileInfoData>(file);
            Marshal.FreeHGlobal(file);
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct FileInfoData
    {
        public uint Size;
        public string Path;
        public IntPtr File;
        public IntPtr Subject;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct TrustData
    {
        public uint Size;
        public IntPtr PolicyCallbackData;
        public IntPtr SipClientData;
        public uint UiChoice;
        public uint RevocationChecks;
        public uint UnionChoice;
        public IntPtr File;
        public uint StateAction;
        public IntPtr StateData;
        public IntPtr UrlReference;
        public uint ProviderFlags;
        public uint UiContext;
        public IntPtr SignatureSettings;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct TrustCertificate
    {
        public uint Size;
        public IntPtr Context;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct CertificateContext
    {
        public uint EncodingType;
        public IntPtr Encoded;
        public uint EncodedLength;
    }

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref TrustData data);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperProvDataFromStateData(IntPtr stateData);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvSignerFromChain(IntPtr providerData, uint index,
        [MarshalAs(UnmanagedType.Bool)] bool counterSigner, uint counterIndex);

    [DllImport("wintrust.dll", ExactSpelling = true)]
    private static extern IntPtr WTHelperGetProvCertFromChain(IntPtr signer, uint index);
}
