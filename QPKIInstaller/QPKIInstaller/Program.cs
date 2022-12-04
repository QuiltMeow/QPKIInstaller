using QPKIInstaller.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;

namespace QPKIInstaller
{
    public static class Program
    {
        private static readonly IDictionary<string, byte[]> ROOT_CERTIFICATE = new Dictionary<string, byte[]>() {
            { "Quilt Root Certification Authority.crt", Resources.Quilt_Root_Certification_Authority }
        };

        private static readonly IDictionary<string, byte[]> CA_CERTIFICATE = new Dictionary<string, byte[]>() {
            { "Quilt Certification Authority.crt", Resources.Quilt_Certification_Authority },
            { "Quilt Organization Certification Authority.crt", Resources.Quilt_Organization_Certification_Authority },
            { "Quilt Personal Certification Authority.crt", Resources.Quilt_Personal_Certification_Authority },
            { "Quilt Server Certification Authority.crt", Resources.Quilt_Server_Certification_Authority },
            { "Quilt Third Party Server Certification Authority.crt", Resources.Quilt_Third_Party_Server_Certification_Authority }
        };

        private static string temporaryDirectory;
        private static readonly Random random = new Random();

        [STAThread]
        public static void Main()
        {
            if (MessageBox.Show("請問是否進行憑證安裝 ?", "QPKI 憑證安裝", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    temporaryDirectory = getUniqueTempFolder();
                    Directory.CreateDirectory(temporaryDirectory);

                    installCertificate();
                    MessageBox.Show("安裝完成", "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"安裝憑證時發生例外狀況 : {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (temporaryDirectory != null)
                    {
                        try
                        {
                            Directory.Delete(temporaryDirectory, true);
                        }
                        catch
                        {
                        }
                    }
                }
            }
        }

        private static string getUniqueTempFolder()
        {
            const int tryTime = 1000;
            const string prefix = "QPKI$";

            for (int i = 1; i <= tryTime; ++i)
            {
                string ret = Path.Combine(Path.GetTempPath(), prefix + random.Next());
                if (!Directory.Exists(ret) && !File.Exists(ret))
                {
                    return ret;
                }
            }
            throw new Exception("無法建立臨時資料夾");
        }

        private static string extractFile(string fileName, byte[] data)
        {
            string path = Path.Combine(temporaryDirectory, fileName);
            File.WriteAllBytes(path, data);
            return path;
        }

        private static void installCertificate(X509Store openStore, IDictionary<string, byte[]> certificateDictionary)
        {
            foreach (KeyValuePair<string, byte[]> pair in certificateDictionary)
            {
                string path = extractFile(pair.Key, pair.Value);
                openStore.Add(new X509Certificate2(X509Certificate.CreateFromCertFile(path)));
            }
        }

        private static void installCertificate()
        {
            using (X509Store store = new X509Store(StoreName.Root, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadWrite);
                installCertificate(store, ROOT_CERTIFICATE);
            }

            using (X509Store store = new X509Store(StoreName.CertificateAuthority, StoreLocation.LocalMachine))
            {
                store.Open(OpenFlags.ReadWrite);
                installCertificate(store, CA_CERTIFICATE);
            }
        }
    }
}