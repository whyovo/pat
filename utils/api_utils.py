import json
import requests
import openai
from urllib.parse import urlparse
import re
import time


class ApiAdapter:
    """API适配器基类"""

    def __init__(self):
        self.chunk_callback = None

    def set_chunk_callback(self, callback):
        """设置接收内容块的回调函数"""
        self.chunk_callback = callback

    def create_completion(
        self,
        prompt=None,
        system=None,
        stream=False,
        timeout=None,
        terminate_check_fn=None,
        messages=None,
    ):
        """
        创建完成请求
        :param prompt: 用户提示词
        :param system: 系统提示词
        :param stream: 是否使用流式输出
        :param timeout: 超时时间（秒）
        :param terminate_check_fn: 终止检查函数
        :param messages: 如果提供，直接使用这些消息而不是构建新的
        :return: 完成结果
        """
        if messages is None:
            if prompt is not None:  # 允许prompt为空
                if system:
                    messages = [
                        {"role": "system", "content": system},
                        {"role": "user", "content": prompt},
                    ]
                else:
                    messages = [{"role": "user", "content": prompt}]
            else:
                # 如果没有提供prompt和messages，则报错
                if not messages:
                    raise ValueError("必须提供prompt或messages参数")

        # 添加日志，记录消息长度
        total_length = sum(len(msg.get("content", "")) for msg in messages)
        print(
            f"API请求: 消息数量={len(messages)}, 总字符数={total_length}, 流式模式={stream}"
        )

        return self._perform_completion(messages, stream, timeout, terminate_check_fn)

    def _perform_completion(
        self, messages, stream=False, timeout=None, terminate_check_fn=None
    ):
        """执行实际的完成请求，由子类实现"""
        raise NotImplementedError("请在子类中实现_perform_completion方法")


class OpenAIAdapter(ApiAdapter):
    """OpenAI API适配器"""

    def __init__(self, api_key, model="gpt-3.5-turbo", api_base=None):
        super().__init__()
        self.model = model
        self.api_key = api_key
        print(
            f"初始化OpenAIAdapter - 模型: {model}, API基础URL: {api_base or 'default'}"
        )
        try:
            # 设置更长的默认超时
            self.client = openai.OpenAI(
                api_key=api_key, base_url=api_base, timeout=120.0  # 默认120秒超时
            )
            print("OpenAI客户端初始化成功")
        except Exception as e:
            print(f"OpenAI客户端初始化失败: {type(e).__name__}: {str(e)}")
            raise

    def _perform_completion(
        self, messages, stream=False, timeout=None, terminate_check_fn=None
    ):
        try:
            # 使用流式响应
            if stream:
                full_content = ""
                start_time = time.time()
                print(
                    f"开始OpenAI流式请求 - 模型: {self.model}, 超时: {timeout or 60}秒"
                )

                # 增加超时时间
                actual_timeout = timeout if timeout else 120

                stream = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    stream=True,
                    timeout=actual_timeout,
                )

                chunk_count = 0
                for chunk in stream:
                    # 检查是否应该终止处理
                    if terminate_check_fn and terminate_check_fn():
                        print("API流被用户终止")
                        return full_content

                    if (
                        chunk.choices
                        and hasattr(chunk.choices[0], "delta")
                        and hasattr(chunk.choices[0].delta, "content")
                    ):
                        content = chunk.choices[0].delta.content
                        if content:
                            full_content += content
                            # 正确调用回调函数
                            if self.chunk_callback:
                                self.chunk_callback(content)
                            # 提供生成器接口
                            chunk_count += 1
                            yield chunk

                duration = time.time() - start_time
                print(
                    f"流式响应完成 - 用时: {duration:.2f}秒, 接收块数: {chunk_count}, 内容长度: {len(full_content)}"
                )
                return full_content
            else:
                # 使用普通响应
                print(f"开始OpenAI非流式请求 - 模型: {self.model}")
                actual_timeout = timeout if timeout else 120

                response = self.client.chat.completions.create(
                    model=self.model, messages=messages, timeout=actual_timeout
                )
                result = response.choices[0].message.content
                print(f"非流式响应完成 - 内容长度: {len(result)}")
                return result
        except Exception as e:
            print(f"OpenAI API调用失败: {type(e).__name__}: {str(e)}")
            raise


class AzureOpenAIAdapter(ApiAdapter):
    """Azure OpenAI API适配器"""

    def __init__(self, api_key, model, api_base):
        super().__init__()
        self.model = model
        self.api_key = api_key
        self.api_base = api_base
        try:
            self.client = openai.AzureOpenAI(
                api_key=api_key, api_version="2023-05-15", azure_endpoint=api_base
            )
        except Exception as e:
            print(f"初始化Azure OpenAI客户端失败: {str(e)}")
            raise

    def _perform_completion(
        self, messages, stream=False, timeout=None, terminate_check_fn=None
    ):
        try:
            # 使用流式响应
            if stream:
                full_content = ""
                stream = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    stream=True,
                    timeout=timeout if timeout else 60,
                )
                for chunk in stream:
                    # 检查是否应该终止处理
                    if terminate_check_fn and terminate_check_fn():
                        return full_content

                    if (
                        chunk.choices
                        and hasattr(chunk.choices[0], "delta")
                        and hasattr(chunk.choices[0].delta, "content")
                    ):
                        content = chunk.choices[0].delta.content
                        if content:
                            full_content += content
                            # 正确调用回调函数
                            if self.chunk_callback:
                                self.chunk_callback(content)
                            # 提供生成器接口
                            yield chunk
                return full_content
            else:
                # 使用普通响应
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    timeout=timeout if timeout else 60,
                )
                return response.choices[0].message.content
        except Exception as e:
            print(f"Azure OpenAI API调用失败: {str(e)}")
            raise


class GenericAPIAdapter(ApiAdapter):
    """通用API适配器，支持兼容OpenAI API的其他服务"""

    def __init__(self, api_key, model, api_base):
        super().__init__()
        self.model = model
        self.api_key = api_key
        self.api_base = api_base

    def _perform_completion(
        self, messages, stream=False, timeout=None, terminate_check_fn=None
    ):
        try:
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}",
            }

            # 构建endpoint，确保以/v1/chat/completions结尾
            endpoint = self.api_base
            if not endpoint.endswith("/v1/chat/completions"):
                if not endpoint.endswith("/"):
                    endpoint += "/"
                endpoint += "v1/chat/completions"

            # 如果需要流式响应
            if stream:
                payload = {
                    "model": self.model,
                    "messages": messages,
                    "stream": True,
                }

                # 发起流式请求
                actual_timeout = timeout if timeout else 120  # 使用更长的超时

                print(
                    f"发起GenericAPI流式请求 - 端点: {endpoint}, 超时: {actual_timeout}秒"
                )

                response = requests.post(
                    endpoint,
                    headers=headers,
                    json=payload,
                    stream=True,
                    timeout=actual_timeout,
                )
                response.raise_for_status()

                full_content = ""
                # 处理SSE流
                for line in response.iter_lines():
                    # 检查是否应该终止处理
                    if terminate_check_fn and terminate_check_fn():
                        return full_content

                    if line:
                        line = line.decode("utf-8")
                        if line.startswith("data: "):
                            if line.strip() == "data: [DONE]":
                                break
                            try:
                                data = json.loads(line[6:])
                                if "choices" in data and len(data["choices"]) > 0:
                                    if (
                                        "delta" in data["choices"][0]
                                        and "content" in data["choices"][0]["delta"]
                                    ):
                                        content = data["choices"][0]["delta"]["content"]
                                        if content:
                                            full_content += content
                                            # 正确调用回调函数
                                            if self.chunk_callback:
                                                self.chunk_callback(content)
                                            # 构造类似openai格式的chunk对象用于生成器
                                            yield data
                            except Exception as e:
                                print(f"处理流数据时出错: {str(e)}")
                return full_content

            else:
                # 普通响应
                payload = {
                    "model": self.model,
                    "messages": messages,
                }

                actual_timeout = timeout if timeout else 120  # 使用更长的超时

                print(
                    f"发起GenericAPI非流式请求 - 端点: {endpoint}, 超时: {actual_timeout}秒"
                )

                response = requests.post(
                    endpoint, headers=headers, json=payload, timeout=actual_timeout
                )
                response.raise_for_status()

                response_data = response.json()
                return response_data["choices"][0]["message"]["content"]

        except Exception as e:
            print(f"通用API调用失败: {type(e).__name__}: {str(e)}")
            raise


def get_api_adapter(api_url=None, api_key=None, model=None):
    """
    根据URL和模型名称选择合适的API适配器
    支持旧版调用方式(model, api_url, api_key)和新版调用方式(api_url, api_key, model)
    """
    # 检测调用方式，兼容旧版接口
    if api_url is not None and api_key is None and model is None:
        # 旧版调用顺序: (model, api_url, api_key)
        model, api_url, api_key = api_url, api_key, model

    # 添加调试信息，显示实际使用的参数
    print(
        f"get_api_adapter - 实际参数: URL='{api_url}', API Key='{api_key[:5]}...(隐藏)', Model='{model}'"
    )

    # 确保参数不为空
    if not api_url:
        print("错误: API URL为空")
        return None

    if not api_key:
        print("错误: API Key为空")
        return None

    try:
        model = model or "gpt-3.5-turbo"

        # 修正URL中的常见问题
        # 1. 移除可能的前后空格
        api_url = api_url.strip()

        # 2. 确保URL以http或https开头
        if not api_url.startswith(("http://", "https://")):
            api_url = "https://" + api_url
            print(f"URL修正: 添加https:// 前缀 -> {api_url}")

        # 3. 修复特定模型的URL问题
        if "deepseek" in model.lower() and "api.deepseek.com" not in api_url.lower():
            # 如果模型包含deepseek但URL不是官方API，使用正确的URL
            api_url = "https://api.deepseek.com"
            print(f"URL修正: 将URL更改为Deepseek官方API -> {api_url}")

        # 解析URL以确定适配器类型
        url_parts = urlparse(api_url)

        # 判断是否是Azure OpenAI
        if "azure.com" in url_parts.netloc:
            print(f"使用Azure OpenAI适配器, 模型: {model}")
            return AzureOpenAIAdapter(api_key, model, api_url)
        # 判断是否是官方OpenAI
        elif "openai.com" in url_parts.netloc or not url_parts.netloc:
            # 官方API或默认
            print(f"使用OpenAI适配器, 模型: {model}")
            return OpenAIAdapter(api_key, model, api_url if url_parts.netloc else None)
        # 判断是否是Deepseek
        elif "deepseek.com" in url_parts.netloc:
            print(f"使用Deepseek适配器, 模型: {model}")
            return OpenAIAdapter(api_key, model, api_url)
        else:
            # 第三方API
            print(f"使用通用API适配器, 模型: {model}, URL: {api_url}")
            return GenericAPIAdapter(api_key, model, api_url)
    except Exception as e:
        print(f"创建API适配器失败: {type(e).__name__}: {str(e)}")
        import traceback

        traceback.print_exc()
        return None


def construct_prompt(system_prompt, user_prompt):
    """构建API请求的消息格式"""
    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]


def clean_text_for_api(text):
    """清理文本使其适合API处理"""
    if not text:
        return ""
    try:
        # 初始化一个空字符串，用于存储清理后的文本
        cleaned_text = ""
        # 遍历输入文本中的每个字符
        for char in text:
            # 检查字符是否是代理对字符（surrogate pairs），范围为 0xD800 到 0xDFFF
            if not (0xD800 <= ord(char) <= 0xDFFF):
                # 如果不是代理对字符，将其添加到清理后的文本中
                cleaned_text += char
            else:
                # 如果是代理对字符，用空格替换
                cleaned_text += " "

        # 再次清理文本，移除 Unicode 范围超过 65536 的字符
        cleaned_text = "".join(
            char if ord(char) < 65536 else " " for char in cleaned_text
        )

        # 使用正则表达式移除控制字符（ASCII 范围内的不可打印字符）
        cleaned_text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", cleaned_text)

        # 返回清理后的文本
        return cleaned_text

    except Exception:
        # 如果在清理过程中发生异常，返回原始文本的 ASCII 编码版本，忽略非 ASCII 字符
        return text.encode("ascii", "ignore").decode("ascii")
