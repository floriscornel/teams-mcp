const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

/**
 * CredentialStore - Secure credential storage with AES-256-CBC encryption
 *
 * CRITICAL: Node.js 22 compatible implementation using createCipheriv with explicit IV
 * The deprecated createCipher() is NOT used - we use createCipheriv() with a random IV
 * for each encryption operation, ensuring proper security.
 *
 * Format: iv:encrypted_data (hex encoded)
 */
class CredentialStore {
    constructor(encryptionKey, storePath = './credentials.json') {
        if (!encryptionKey || encryptionKey.length < 32) {
            throw new Error('Encryption key must be at least 32 characters long for AES-256');
        }

        // Derive a 32-byte key from the encryption key using SHA-256
        this.key = crypto.createHash('sha256').update(encryptionKey).digest();
        this.storePath = path.resolve(storePath);

        // Ensure the credentials file exists
        this._ensureStoreExists();

        console.log(`[CredentialStore] Initialized with store at: ${this.storePath}`);
    }

    /**
     * Ensure the credentials store file exists
     * @private
     */
    _ensureStoreExists() {
        try {
            if (!fs.existsSync(this.storePath)) {
                fs.writeFileSync(this.storePath, JSON.stringify({}), 'utf8');
                console.log(`[CredentialStore] Created new credentials store`);
            }
        } catch (error) {
            console.error(`[CredentialStore] Error creating store:`, error);
            throw new Error(`Failed to initialize credential store: ${error.message}`);
        }
    }

    /**
     * Encrypt text using AES-256-CBC with a random IV
     * @param {string} text - Plain text to encrypt
     * @returns {string} Encrypted text in format "iv:encrypted" (hex encoded)
     */
    encrypt(text) {
        try {
            // Generate a random 16-byte IV for this encryption
            const iv = crypto.randomBytes(16);

            // Create cipher with explicit IV
            const cipher = crypto.createCipheriv('aes-256-cbc', this.key, iv);

            // Encrypt the text
            let encrypted = cipher.update(text, 'utf8', 'hex');
            encrypted += cipher.final('hex');

            // Return IV:encrypted format
            return iv.toString('hex') + ':' + encrypted;
        } catch (error) {
            console.error(`[CredentialStore] Encryption error:`, error);
            throw new Error(`Encryption failed: ${error.message}`);
        }
    }

    /**
     * Decrypt text encrypted with AES-256-CBC
     * @param {string} encryptedText - Encrypted text in format "iv:encrypted"
     * @returns {string} Decrypted plain text
     */
    decrypt(encryptedText) {
        try {
            // Split IV and encrypted data
            const parts = encryptedText.split(':');
            if (parts.length !== 2) {
                throw new Error('Invalid encrypted text format. Expected "iv:encrypted"');
            }

            const iv = Buffer.from(parts[0], 'hex');
            const encrypted = parts[1];

            // Create decipher with the extracted IV
            const decipher = crypto.createDecipheriv('aes-256-cbc', this.key, iv);

            // Decrypt the text
            let decrypted = decipher.update(encrypted, 'hex', 'utf8');
            decrypted += decipher.final('utf8');

            return decrypted;
        } catch (error) {
            console.error(`[CredentialStore] Decryption error:`, error);
            throw new Error(`Decryption failed: ${error.message}`);
        }
    }

    /**
     * Load credentials from disk
     * @returns {Object} Credentials object
     * @private
     */
    _loadStore() {
        try {
            const data = fs.readFileSync(this.storePath, 'utf8');
            return JSON.parse(data);
        } catch (error) {
            console.error(`[CredentialStore] Error loading store:`, error);
            return {};
        }
    }

    /**
     * Save credentials to disk
     * @param {Object} store - Credentials object to save
     * @private
     */
    _saveStore(store) {
        try {
            fs.writeFileSync(this.storePath, JSON.stringify(store, null, 2), 'utf8');
        } catch (error) {
            console.error(`[CredentialStore] Error saving store:`, error);
            throw new Error(`Failed to save credentials: ${error.message}`);
        }
    }

    /**
     * Store encrypted credentials for a user
     * @param {string} userId - User identifier (Teams user ID)
     * @param {string} username - OData service username
     * @param {string} password - OData service password
     */
    storeCredentials(userId, username, password) {
        try {
            console.log(`[CredentialStore] Storing credentials for user: ${userId}`);

            const store = this._loadStore();

            // Encrypt the credentials
            const encryptedUsername = this.encrypt(username);
            const encryptedPassword = this.encrypt(password);

            // Store with timestamp
            store[userId] = {
                username: encryptedUsername,
                password: encryptedPassword,
                updatedAt: new Date().toISOString()
            };

            this._saveStore(store);
            console.log(`[CredentialStore] Credentials stored successfully for user: ${userId}`);

            return true;
        } catch (error) {
            console.error(`[CredentialStore] Error storing credentials:`, error);
            throw error;
        }
    }

    /**
     * Retrieve and decrypt credentials for a user
     * @param {string} userId - User identifier (Teams user ID)
     * @returns {Object|null} Object with {username, password} or null if not found
     */
    getCredentials(userId) {
        try {
            const store = this._loadStore();

            if (!store[userId]) {
                console.log(`[CredentialStore] No credentials found for user: ${userId}`);
                return null;
            }

            // Decrypt the credentials
            const username = this.decrypt(store[userId].username);
            const password = this.decrypt(store[userId].password);

            console.log(`[CredentialStore] Retrieved credentials for user: ${userId}`);

            return { username, password };
        } catch (error) {
            console.error(`[CredentialStore] Error retrieving credentials:`, error);
            return null;
        }
    }

    /**
     * Check if credentials exist for a user
     * @param {string} userId - User identifier
     * @returns {boolean} True if credentials exist
     */
    hasCredentials(userId) {
        const store = this._loadStore();
        return !!store[userId];
    }

    /**
     * Delete credentials for a user
     * @param {string} userId - User identifier
     * @returns {boolean} True if credentials were deleted
     */
    deleteCredentials(userId) {
        try {
            console.log(`[CredentialStore] Deleting credentials for user: ${userId}`);

            const store = this._loadStore();

            if (!store[userId]) {
                return false;
            }

            delete store[userId];
            this._saveStore(store);

            console.log(`[CredentialStore] Credentials deleted for user: ${userId}`);
            return true;
        } catch (error) {
            console.error(`[CredentialStore] Error deleting credentials:`, error);
            return false;
        }
    }

    /**
     * Get all user IDs that have stored credentials
     * @returns {string[]} Array of user IDs
     */
    getAllUserIds() {
        const store = this._loadStore();
        return Object.keys(store);
    }

    /**
     * Clear all credentials (use with caution!)
     * @returns {boolean} True if successful
     */
    clearAll() {
        try {
            console.warn(`[CredentialStore] Clearing ALL credentials!`);
            this._saveStore({});
            return true;
        } catch (error) {
            console.error(`[CredentialStore] Error clearing credentials:`, error);
            return false;
        }
    }
}

module.exports = CredentialStore;
