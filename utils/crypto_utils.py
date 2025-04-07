"""
实现配置加密和解密功能
"""

import os
import base64
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend


def _get_key_iv(salt=None):
    """生成加密密钥和IV"""
    # 使用固定密钥作为基础（在实际应用中应当使用更安全的方法）
    base_key = b"PaperAnalyzerSecretKey2023"

    if salt is None:
        salt = os.urandom(16)

    # 使用PBKDF2派生实际密钥
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
        backend=default_backend(),
    )
    key = kdf.derive(base_key)

    # 使用密钥的一部分作为IV
    iv = key[:16]

    return key, iv, salt


def encrypt_config(config):
    """加密配置信息，主要保护API密钥"""
    # 只有API密钥需要加密
    if "key" in config and config["key"]:
        try:
            api_key = config["key"].encode("utf-8")
            salt = os.urandom(16)
            key, iv, _ = _get_key_iv(salt)

            # 使用AES-CBC加密
            cipher = Cipher(
                algorithms.AES(key), modes.CBC(iv), backend=default_backend()
            )
            encryptor = cipher.encryptor()

            # PKCS7填充
            block_size = 16
            padding_len = block_size - (len(api_key) % block_size)
            api_key += bytes([padding_len]) * padding_len

            encrypted_key = encryptor.update(api_key) + encryptor.finalize()

            # 将加密后的内容和salt组合并使用base64编码
            encrypted_data = salt + encrypted_key
            config["key"] = base64.b64encode(encrypted_data).decode("utf-8")
            config["encrypted"] = True

            return config
        except Exception as e:
            print(f"加密失败: {str(e)}")

    return config


def decrypt_config(config):
    """解密配置信息，恢复API密钥"""
    if config.get("encrypted", False) and "key" in config and config["key"]:
        try:
            encrypted_data = base64.b64decode(config["key"])
            salt = encrypted_data[:16]
            encrypted_key = encrypted_data[16:]

            key, iv, _ = _get_key_iv(salt)

            # 使用AES-CBC解密
            cipher = Cipher(
                algorithms.AES(key), modes.CBC(iv), backend=default_backend()
            )
            decryptor = cipher.decryptor()

            decrypted_data = decryptor.update(encrypted_key) + decryptor.finalize()

            # 移除PKCS7填充
            padding_len = decrypted_data[-1]
            api_key = decrypted_data[:-padding_len].decode("utf-8")

            config["key"] = api_key
            config["encrypted"] = False

            return config
        except Exception as e:
            print(f"解密失败: {str(e)}")
            # 解密失败时返回空密钥
            config["key"] = ""
            config["encrypted"] = False

    return config
