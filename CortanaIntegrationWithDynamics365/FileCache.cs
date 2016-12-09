using System;
using System.Security.Cryptography;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.IO;
using Windows.Storage;
using Windows.Security.Cryptography;
using Windows.Storage.Streams;

namespace CortanaIntegrationWithDynamics365
{
    class FileCache : TokenCache
    {
        public string CacheFilePath;
        private static readonly object FileLock = new object();
        public StorageFolder appFolder = ApplicationData.Current.LocalFolder;
        // Initializes the cache against a local file.
        // If the file is already present, it loads its content in the ADAL cache
        public FileCache(string filePath = "\\TokenCache.dat")
        {
            CacheFilePath = appFolder.Path + filePath;
            this.AfterAccess = AfterAccessNotification;
            this.BeforeAccess = BeforeAccessNotification;
            lock (FileLock)
            {
                this.Deserialize(File.Exists(CacheFilePath) ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), null, DataProtectionScope.CurrentUser) : null);
            }
        }

        // Empties the persistent store.
        public override void Clear()
        {
            base.Clear();
            File.Delete(CacheFilePath);
        }

        // Triggered right before ADAL needs to access the cache.
        // Reload the cache from the persistent store in case it changed since the last access.
        void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            lock (FileLock)
            {
                this.Deserialize(File.Exists(CacheFilePath) ? ProtectedData.Unprotect(File.ReadAllBytes(CacheFilePath), null, DataProtectionScope.CurrentUser) : null);
            }
        }

        // Triggered right after ADAL accessed the cache.
        async void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (this.HasStateChanged)
            {
                StorageFile cacheFile = await appFolder.CreateFileAsync("TokenCache.dat", CreationCollisionOption.ReplaceExisting);
                await FileIO.WriteBufferAsync(cacheFile, CryptographicBuffer.CreateFromByteArray(ProtectedData.Protect(this.Serialize(), null, DataProtectionScope.CurrentUser)));
                this.HasStateChanged = false;
            }
        }
    }
}