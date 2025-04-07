import os
import json
from utils.crypto_utils import encrypt_config, decrypt_config

def get_config_path():
    """获取API配置文件路径"""
    user_docs = os.path.join(os.path.expanduser('~'), 'Documents', '论文分析工具')
    os.makedirs(user_docs, exist_ok=True)
    return os.path.join(user_docs, 'api_config.json')

def load_api_configs():
    """加载API配置，使用加密保护API密钥"""
    try:
        config_path = get_config_path()
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 检查是否加密
                if config.get("encrypted", False):
                    # 解密配置
                    config = decrypt_config(config)
                return config
        return {"url": "", "key": "", "model": "", "remember": True, "encrypted": False}
    except Exception as e:
        print(f"加载API配置时出错: {str(e)}")
        return {"url": "", "key": "", "model": "", "remember": True, "encrypted": False}

def save_api_configs(api_url, api_key, api_model, remember_key):
    """保存API配置到文件，使用加密保护API密钥"""
    try:
        # 只有当"记住API设置"被勾选时才保存
        if not remember_key:
            return
            
        user_docs = os.path.join(os.path.expanduser('~'), 'Documents', '论文分析工具')
        if not os.path.exists(user_docs):
            os.makedirs(user_docs)
            
        config_path = os.path.join(user_docs, 'api_config.json')
        config = {
            "url": api_url,
            "key": api_key,
            "model": api_model,
            "remember": remember_key,
            "encrypted": False  # 初始标记为未加密
        }
        
        # 加密配置
        config = encrypt_config(config)
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
            print("API配置已保存并加密")
    except Exception as e:
        print(f"保存API配置出错: {str(e)}")
