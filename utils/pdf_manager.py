import os
import re
import time
import threading
import pandas as pd
import tkinter as tk
import datetime
from collections import Counter
from tkinter import messagebox, filedialog

from utils.api_utils import get_api_adapter, clean_text_for_api
from utils.app_manager import get_thread_safe_gui
from utils.thread_utils import ThreadSafeText


def analyze_papers(self):
    """开始分析论文"""
    if not hasattr(self, "pdf_paths") or not self.pdf_paths:
        self.status_bar["text"] = "错误：请先选择PDF文件"
        messagebox.showerror("错误", "请先选择要分析的PDF文件")
        return

    if not self.excel_path.get():
        self.status_bar["text"] = "错误：请先选择或创建Excel文件"
        messagebox.showerror("错误", "请先选择或创建要保存结果的Excel文件")
        return

    self.disable_analysis_buttons()
    # 确保取消按钮可用且显示"取消任务"
    if hasattr(self, "cancel_analysis_btn"):
        self.cancel_analysis_btn.configure(text="取消任务", state="normal")

    self.status_bar["text"] = "准备开始分析..."
    self.output_text.insert(tk.END, "\n开始处理论文分析任务...\n", "info")
    self.output_text.tag_configure("info", foreground=self.colors["fg"])
    self.output_text.see(tk.END)
    self.root.update_idletasks()

    df = pd.read_excel(self.excel_path.get())
    total_files = len(self.pdf_paths)

    api_url, api_key = self.get_api_info()

    analysis_thread = threading.Thread(
        target=self.process_papers_async,
        args=(df, api_url, api_key, total_files),
        daemon=True,
    )
    analysis_thread.start()


CANCELLED = False


def set_cancelled(flag):
    global CANCELLED
    CANCELLED = flag


def process_papers_async(self, df, api_url, api_key, total_files):
    """处理多个PDF文件的异步函数"""
    tm = get_thread_safe_gui(self.root)
    self.cancel_analysis_requested = False
    set_cancelled(False)

    try:
        # 检查API信息
        if not api_key:
            tm.add_task(
                self.show_error_and_reset, "API密钥为空，请在API设置中填写有效的密钥"
            )
            return

        model = self.api_model_var.get() or "gpt-3.5-turbo"
        tm.add_task(self.output_text.insert, tk.END, f"使用模型: {model}\n", "info")

        # 创建API适配器
        api_adapter = get_api_adapter(api_url, api_key, model)
        if not api_adapter:
            tm.add_task(self.show_error_and_reset, "无法创建API适配器，请检查API设置")
            return

        # 初始化结果DataFrame
        results = []

        # 处理每个PDF文件
        for i, pdf_path in enumerate(self.pdf_paths, 1):
            if self.cancel_analysis_requested:
                tm.add_task(
                    self.output_text.insert, tk.END, "\n用户取消了分析任务\n", "warning"
                )
                tm.add_task(
                    self.output_text.tag_configure, "warning", foreground="#e69138"
                )
                break

            tm.add_task(self.update_progress_status, i, total_files)

            # 更新输出文本
            pdf_name = os.path.basename(pdf_path)
            tm.add_task(
                self.output_text.insert,
                tk.END,
                f"\n-- 正在处理 ({i}/{total_files}): {pdf_name} --\n",
                "subheader",
            )
            tm.add_task(
                self.output_text.tag_configure,
                "subheader",
                foreground="#5ba3e0",
                font=("微软雅黑", 10, "bold"),
            )

            # 提取PDF文本
            out = ThreadSafeText(self.output_text, self.root)
            pdf_text = extract_pdf_text(pdf_path, out)

            if not pdf_text or len(pdf_text) < 100:
                out.insert(
                    tk.END,
                    "警告: PDF文本提取失败或文本内容过少，无法进行分析\n",
                    "warning",
                )
                continue

            # 调用API并解析结果
            try:
                out.insert(tk.END, "调用API中，请稍候...\n", "processing")

                response = call_api_with_retry(self, "", pdf_text, out, tm)

                if not response or self.cancel_analysis_requested:
                    if self.cancel_analysis_requested:
                        out.insert(tk.END, "已取消分析\n", "warning")
                    else:
                        out.insert(tk.END, "API响应为空，跳过此文件\n", "error")
                    continue

                # 解析响应
                data = extract_data_from_response(response)

                # 检查提取的数据
                if not data.get("论文年份"):
                    # 尝试从文件名中提取年份
                    year_match = re.search(r"(19|20)\d{2}", pdf_name)
                    if year_match:
                        data["论文年份"] = year_match.group(0)
                    else:
                        # 默认使用当前年份
                        data["论文年份"] = str(datetime.datetime.now().year - 1)

                # 添加结果
                results.append(data)

                # 输出摘要部分 - 修改预览文本长度并确保完整显示
                tm.add_task(
                    self.output_text.insert, tk.END, "\n摘要预览:\n", "preview_header"
                )
                tm.add_task(
                    self.output_text.tag_configure,
                    "preview_header",
                    foreground="#4a8cca",
                    font=("微软雅黑", 10, "bold"),
                )

                # 修改这里：增加预览摘要的长度并用省略号表示截断
                preview_text = data.get("论文摘要（中文）", "未能提取摘要")
                if len(preview_text) > 300:  # 从150扩展到300
                    preview_text = preview_text[:300] + "..."

                tm.add_task(
                    self.output_text.insert, tk.END, f"{preview_text}\n", "preview_text"
                )

                # 显示完成标记
                tm.add_task(self.output_text.insert, tk.END, "✓ 分析完成\n", "success")
                tm.add_task(
                    self.output_text.tag_configure, "success", foreground="#8bc34a"
                )

            except Exception as api_error:
                out.insert(tk.END, f"API调用或解析出错: {str(api_error)}\n", "error")
                tm.add_task(
                    self.output_text.tag_configure, "error", foreground="#ef5350"
                )
                continue

        # 保存结果到Excel
        if results and not self.cancel_analysis_requested:
            result_df = pd.DataFrame(results)
            tm.add_task(
                self.output_text.insert,
                tk.END,
                "\n所有PDF处理完成，正在保存到Excel...\n",
                "info",
            )
            tm.add_task(self.save_excel_result, result_df)
        elif self.cancel_analysis_requested:
            tm.add_task(
                self.output_text.insert, tk.END, "\n分析被取消，未保存结果\n", "warning"
            )
        else:
            tm.add_task(
                self.output_text.insert,
                tk.END,
                "\n没有成功分析的PDF，未保存结果\n",
                "warning",
            )

        # 恢复UI状态
        tm.add_task(self.on_analysis_complete)

    except Exception as e:
        import traceback

        print(f"PDF处理异常: {str(e)}")
        print(traceback.format_exc())
        tm.add_task(self.show_error_and_reset, f"处理过程中发生错误: {str(e)}")


def extract_pdf_text(path, out):
    """从PDF提取文本"""
    out.insert(tk.END, "正在提取PDF文本...\n")

    try:
        import PyPDF2

        with open(path, "rb") as f:
            pdf_reader = PyPDF2.PdfReader(f)
            num_pages = len(pdf_reader.pages)

            # 提取前30页或所有页面（取较小值）
            max_pages = min(30, num_pages)
            text_parts = []

            for i in range(max_pages):
                page = pdf_reader.pages[i]
                text = page.extract_text()
                if text:
                    text_parts.append(text)

            combined_text = "\n".join(text_parts)

            # 进行基本清理
            combined_text = re.sub(r"\s+", " ", combined_text)  # 合并多余空白
            combined_text = re.sub(r"\n+", "\n", combined_text)  # 合并多余换行

            # 添加API清理步骤，确保文本适合API处理
            combined_text = clean_text_for_api(combined_text)

            extracted_percent = (max_pages / num_pages) * 100
            out.insert(
                tk.END,
                f"已提取 {max_pages}/{num_pages} 页 ({extracted_percent:.1f}%)\n",
            )
            return combined_text

    except Exception as e:
        out.insert(tk.END, f"PDF文本提取失败: {str(e)}\n", "error")
        return ""


def call_api_with_retry(self, system_prompt, prompt, out, tm):
    """调用API并支持重试机制"""
    api_url, api_key = self.get_api_info()
    model = self.api_model_var.get() or "gpt-3.5-turbo"

    # 增加调试日志
    print(f"API调用开始 - URL: {api_url}, 模型: {model}, 文本长度: {len(prompt)}")
    out.insert(tk.END, f"使用模型: {model}\n", "info")

    # 使用正确的提示词构造方式
    try:
        from configs.excel_header_config import load_custom_columns

        custom_columns = load_custom_columns()
        # 根据自定义表头动态构建提示词格式
        format_str = ""
        for column in custom_columns:
            format_str += f"{column}|[内容]\n"
        print(f"已加载自定义表头, 共{len(custom_columns)}列")
    except Exception as e:
        print(f"加载自定义表头出错: {str(e)}")
        # 使用安全的默认表头格式 - 只包含最基本的必需字段
        format_str = """论文年份|[年份]
论文英文引用信息|[引用信息]
论文摘要（中文）|[摘要内容]
研究问题及创新点|[内容]
研究意义|[内容]
研究对象及特点|[内容]
实验设置|[内容]
数据分析方法|[内容]
重要结论|[内容]
未来研究展望|[内容]"""
        print("使用默认表头格式")

    # 构造API提示词
    system_prompt = "你是一位学术研究总结专家，擅长对各类学术论文进行深入分析和总结。"

    # 清理PDF文本确保适合API处理
    cleaned_text = clean_text_for_api(prompt)

    # 限制文本长度，避免超出模型上下文窗口
    max_length = 60000  # 根据模型调整，GPT-3.5可能需要较短
    if len(cleaned_text) > max_length:
        out.insert(
            tk.END,
            f"原文本长度: {len(cleaned_text)}字符，截断到{max_length}字符\n",
            "warning",
        )
        cleaned_text = cleaned_text[:max_length] + "..."

    print(f"清理后文本长度: {len(cleaned_text)}")

    user_prompt = f"""
\"Background\": \"用户需要对学术论文进行详细的总结，用于文献综述或研究整理。\",
\"Skills\": \"你具备强大的文献阅读和分析能力，能够快速理解论文的核心内容。\",
\"Goals\": \"根据用户提供的论文信息，详细总结论文的各个方面，并严格按照以下格式提供分析结果（确保使用|分隔符，并且只输出下面指定的字段）：
{format_str}\",
\"Constrains\": \"内容应符合学术规范，信息准确、完整，且具有一定的逻辑性和可读性。必须严格按照指定格式返回结果。只返回指定字段，不要添加额外字段。如果需要回复研究意义，确保只回复一段话且可以直接运用于综述。\",
\"论文内容\": \"{cleaned_text}\"
"""

    max_retries = 3
    retry_delay = 5  # 秒

    # 添加响应验证函数
    def is_valid_response(text):
        """检查API响应是否有效"""
        if not text or not text.strip():
            return False

        # 检查是否包含至少一对键值对（使用|分隔）
        if "|" not in text:
            # 尝试识别其他可能的分隔符（如果API响应格式不正确）
            if "：" in text or ":" in text:
                print("检测到使用了非标准分隔符（：或:），尝试处理...")
                return True  # 允许使用其他分隔符的响应通过

            # 检查是否包含关键字段名
            key_fields = ["论文年份", "论文摘要", "研究问题", "研究意义", "结论"]
            for field in key_fields:
                if field in text:
                    print(f"检测到关键字段：{field}，尝试处理...")
                    return True

            print("未能在响应中检测到有效格式或关键字段")
            return False

        # 默认认为响应有效
        return True

    # 用于收集和存储流式响应的变量
    complete_response_text = ""
    broken_chunk = False
    api_adapters_tried = []

    for attempt in range(max_retries):
        try:
            # 检查是否取消
            if self.cancel_analysis_requested:
                print("用户取消API调用")
                return None

            # 创建API适配器
            print(f"尝试创建API适配器 (尝试 {attempt+1}/{max_retries})")

            # 确保我们不会重复使用失败的适配器
            api_adapter = None
            while api_adapter is None:
                api_adapter = get_api_adapter(api_url, api_key, model)
                current_adapter_info = f"{type(api_adapter).__name__}_{model}"

                if api_adapter is None:
                    print("API适配器创建失败，重试...")
                elif current_adapter_info in api_adapters_tried:
                    # 如果已经尝试过这个适配器，尝试切换模型
                    print(f"已尝试过 {current_adapter_info}，尝试其他模型")
                    if "gpt-3.5" in model:
                        model = model.replace("gpt-3.5", "gpt-4")
                    elif "gpt-4" in model:
                        model = model.replace("gpt-4", "gpt-3.5")
                    elif "-chat" in model:
                        model = model.replace("-chat", "")
                    else:
                        model = model + "-turbo" if "-turbo" not in model else model

                    api_adapter = None  # 重新获取适配器
                else:
                    # 记录已尝试的适配器类型
                    api_adapters_tried.append(current_adapter_info)

            if not api_adapter:
                raise Exception("无法创建API适配器")

            print(f"API适配器创建成功：{type(api_adapter).__name__}，开始API调用")
            out.insert(tk.END, "正在向API发送请求...\n", "info")
            out.flush()  # 确保显示最新状态

            # 重置响应文本
            complete_response_text = ""
            received_chunks = 0

            # 设置更长的超时时间
            timeout = 180  # 3分钟

            # 设置流式响应回调
            def on_chunk_received(chunk):
                nonlocal complete_response_text
                nonlocal received_chunks

                if chunk:
                    complete_response_text += chunk
                    received_chunks += 1

                # 写入UI
                out.insert(tk.END, chunk)

            # 设置回调并调用API
            api_adapter.set_chunk_callback(on_chunk_received)
            print("开始API流式请求，超时设置为", timeout, "秒")

            # 暂存响应对象
            response_obj = None

            try:
                # 调用API获取流式响应
                response_obj = api_adapter.create_completion(
                    prompt=user_prompt,
                    system=system_prompt,
                    stream=True,
                    timeout=timeout,
                )

                if not response_obj:
                    print("API返回空响应对象")
                    raise Exception("API返回空响应对象")
            except Exception as api_err:
                print(f"API初始调用失败: {type(api_err).__name__}: {api_err}")
                broken_chunk = True
                raise api_err

            # 处理流式响应
            try:
                chunk_count = 0

                # 确保刷新任何剩余的文本
                out.flush()

                print("开始处理API响应流")
                for chunk in response_obj:
                    # 检查是否应该终止处理
                    if self.cancel_analysis_requested:
                        print("用户取消，终止响应处理")
                        break

                    # 提取内容
                    chunk_content = None
                    if hasattr(chunk, "choices") and len(chunk.choices) > 0:
                        if hasattr(chunk.choices[0], "delta") and hasattr(
                            chunk.choices[0].delta, "content"
                        ):
                            chunk_content = chunk.choices[0].delta.content

                    if chunk_content:
                        chunk_count += 1

                # 收集流式响应后的处理结果
                print(
                    f"API响应流处理完成，收到 {chunk_count} 个块，总长度 {len(complete_response_text)}"
                )

                # 如果通过回调收集的数据为空，但通过遍历收集的不为空，则使用后者
                if not complete_response_text and chunk_count > 0:
                    print("警告：回调收集响应为空，但流式处理收到了内容")
                    broken_chunk = True

                # 确保保存了响应文本
                if received_chunks == 0 and chunk_count == 0:
                    print("未收到任何内容块")

                if complete_response_text.strip():
                    # 打印响应内容的前100个字符作为调试用途
                    preview = complete_response_text[:100].replace("\n", "\\n")
                    print(f"收到的响应内容预览: {preview}...")

                if is_valid_response(complete_response_text):
                    out.insert(tk.END, "\n✅ API响应接收完成\n", "success")
                    out.tag_configure(
                        "success", foreground="#8bc34a", font=("微软雅黑", 10, "bold")
                    )
                    return complete_response_text
                else:
                    print("API响应内容不符合预期格式")
                    raise Exception("API响应格式不正确")

            except Exception as stream_err:
                print(f"处理流式响应时出错: {type(stream_err).__name__}: {stream_err}")
                broken_chunk = True

                raise stream_err

        except Exception as e:
            print(f"API调用异常: {type(e).__name__}: {str(e)}")

            # 如果已经收集了一些内容，并且看起来是有效的，返回它
            if (
                complete_response_text
                and "|" in complete_response_text
                and len(complete_response_text) > 100
            ):
                print("虽然发生错误，但已收集到有用内容，尝试使用")
                out.insert(tk.END, "\n⚠️ API响应部分完成，尝试处理\n", "warning")
                return complete_response_text

            if attempt < max_retries - 1 and not self.cancel_analysis_requested:
                out.insert(
                    tk.END,
                    f"API调用失败 (尝试 {attempt+1}/{max_retries}): {str(e)}\n",
                    "warning",
                )
                out.insert(tk.END, f"等待 {retry_delay} 秒后重试...\n", "info")
                out.flush()  # 确保显示最新状态

                # 等待指定时间，期间检查是否取消
                for i in range(retry_delay):
                    if self.cancel_analysis_requested:
                        print("等待重试期间用户取消")
                        return None
                    time.sleep(1)
                    if i % 2 == 0:  # 每2秒更新一次倒计时
                        remaining = retry_delay - i
                        out.insert(tk.END, f"再等 {remaining} 秒...\r", "info")
            else:
                print("达到最大重试次数或用户取消，放弃API调用")
                out.insert(tk.END, f"API调用失败: {str(e)}\n", "error")

                # 如果有一些内容，即使不符合预期格式，也尝试使用它
                if complete_response_text and len(complete_response_text) > 50:
                    print("返回不完整的响应用于尝试处理")
                    out.insert(tk.END, "尝试处理不完整响应...\n", "warning")
                    return complete_response_text

                return None

    print("所有重试失败，返回None")
    return None


def collect_response_stream(self, response, out, tm):
    """收集API流式响应，添加超时控制和取消检查"""
    full_text = ""
    display_buffer = ""
    buffer_size = 50  # 字符数

    try:
        for chunk in response:
            # 检查取消请求
            if self.cancel_analysis_requested:
                return None

            if content := chunk.choices[0].delta.content:
                full_text += content
                display_buffer += content

                # 当缓冲区达到一定大小时更新UI
                if len(display_buffer) >= buffer_size:
                    self.append_response_chunk(display_buffer)
                    display_buffer = ""

        # 显示剩余的缓冲区内容
        if display_buffer:
            self.append_response_chunk(display_buffer)

        return full_text
    except Exception as e:
        tm.add_task(
            self.output_text.insert, tk.END, f"\n接收响应时出错: {str(e)}\n", "error"
        )
        return full_text if full_text else None


def extract_data_from_response(text):
    """从API响应提取结构化数据"""
    data = {}
    print(f"开始从API响应提取数据，响应长度: {len(text)}")

    # 使用自定义表头加载列
    try:
        from configs.excel_header_config import load_custom_columns

        fields = load_custom_columns()
        print(f"加载了 {len(fields)} 个自定义字段")
    except Exception as e:
        print(f"加载自定义表头出错: {str(e)}")
        # 定义默认需要提取的字段
        fields = [
            "论文年份",
            "论文英文引用信息",
            "论文摘要（中文）",
            "研究问题及创新点",
            "研究意义",
            "研究对象及特点",
            "实验设置",
            "数据分析方法",
            "重要结论",
            "未来研究展望",
        ]
        print(f"使用默认字段列表: {len(fields)} 个字段")

    # 规范化响应文本以防异常格式
    # 有时API可能返回非标准格式，如使用"："而不是"|"作为分隔符
    normalized_text = text
    if "|" not in text and ("：" in text or ":" in text):
        print("检测到非标准分隔符，尝试规范化")
        lines = text.split("\n")
        normalized_lines = []

        for line in lines:
            # 检查这一行是否可能是字段定义
            for field in fields:
                if line.strip().startswith(field) and ("：" in line or ":" in line):
                    # 替换第一个冒号为管道符
                    if "：" in line:
                        normalized_line = line.replace("：", "|", 1)
                    else:
                        normalized_line = line.replace(":", "|", 1)
                    normalized_lines.append(normalized_line)
                    break
            else:
                normalized_lines.append(line)

        normalized_text = "\n".join(normalized_lines)
        print("规范化后的文本：")
        print(
            normalized_text[:200] + "..."
            if len(normalized_text) > 200
            else normalized_text
        )

    # 从响应中提取信息
    lines = normalized_text.split("\n")
    print(f"响应包含 {len(lines)} 行")

    # 检查是否有直接可用的字段-值结构
    extracted_fields = 0

    # 追踪当前正在处理的字段，用于处理多行内容
    current_field = None
    current_value = ""

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 寻找以 | 分隔的行
        if "|" in line:
            # 如果有待处理的当前字段，先保存它
            if current_field and current_field in fields:
                data[current_field] = current_value.strip()
                extracted_fields += 1

            parts = line.split("|", 1)
            if len(parts) == 2:
                current_field = parts[0].strip()
                current_value = parts[1].strip()

                print(f"找到字段: '{current_field}' 值长度: {len(current_value)}")

                # 如果该字段已是最后一个需要的字段，直接保存
                if current_field in fields and current_field == fields[-1]:
                    data[current_field] = current_value
                    extracted_fields += 1
                    current_field = None
        elif current_field:
            # 附加到当前值（处理多行内容）
            current_value += " " + line

    # 处理最后一个字段（如果有）
    if current_field and current_field in fields:
        data[current_field] = current_value.strip()
        extracted_fields += 1

    print(f"成功提取了 {extracted_fields}/{len(fields)} 个字段")

    # 确保所有字段都有值，即使是空值
    for field in fields:
        if field not in data:
            data[field] = "未提供相关信息"
            print(f"字段 '{field}' 未找到，设置为默认值")

    return data


def regenerate(self):
    """重新生成分析"""
    if not hasattr(self, "pdf_paths") or not self.pdf_paths:
        self.status_bar["text"] = "错误：没有可重新生成的论文"
        messagebox.showerror("错误", "没有可重新生成的论文")
        return
    if not hasattr(self, "analysis_completed") or not self.analysis_completed:
        self.status_bar["text"] = "错误：请先完成分析"
        messagebox.showerror("错误", "请先完成分析")
        return

    self.cancel_analysis_requested = False
    self.analysis_completed = False
    self.disable_analysis_buttons()
    self.status_bar["text"] = "准备重新生成..."
    self.root.update_idletasks()

    try:
        if hasattr(self, "regenerate_btn"):
            self.regenerate_btn["state"] = "disabled"

        threading.Thread(target=self.perform_regenerate, daemon=True).start()
    except Exception as e:
        self.status_bar["text"] = f"重新生成出错: {str(e)}"
        self.enable_analysis_buttons(
            keep_regenerate_disabled=True
        )  # 出错时保持重新生成按钮禁用


def perform_regenerate(self):
    """执行重新生成操作"""
    tm = get_thread_safe_gui(self.root)

    try:
        if os.path.exists(self.excel_path.get()):
            # 读取Excel文件
            df = pd.read_excel(self.excel_path.get())
            pdf_filenames = [os.path.basename(path) for path in self.pdf_paths]
            file_count = len(pdf_filenames)

            if len(df) > 0 and file_count > 0:
                original_count = len(df)
                rows_to_keep = []
                header_rows = []
                matched_rows = []

                # 1. 查找表头分隔行和匹配行
                for idx, row in df.iterrows():
                    # 检查是否有表头分隔标记
                    has_header_mark = False
                    for col in df.columns:
                        if (
                            isinstance(row[col], str)
                            and "--以下使用新表头--" in row[col]
                        ):
                            header_rows.append(idx)
                            has_header_mark = True
                            break

                    if has_header_mark:
                        continue

                    # 检查是否有文件名匹配
                    match_found = False
                    for pdf_name in pdf_filenames:
                        base_name = os.path.splitext(pdf_name)[0]
                        pdf_name_lower = pdf_name.lower()
                        base_name_lower = base_name.lower()

                        for col in df.columns:
                            cell_value = str(row[col]).lower()
                            if (
                                pdf_name_lower in cell_value
                                or (
                                    len(base_name) > 4 and base_name_lower in cell_value
                                )
                                or (
                                    len(cell_value) > 10
                                    and cell_value in pdf_name_lower
                                )
                                or (
                                    len(cell_value) > 10
                                    and cell_value in base_name_lower
                                )
                            ):
                                match_found = True
                                matched_rows.append(idx)
                                break

                        if match_found:
                            break

                    # 保留不匹配的行
                    if not match_found:
                        rows_to_keep.append(idx)

                # 2. 根据查找结果处理数据
                if header_rows:
                    last_header = header_rows[-1]
                    df = df.iloc[: last_header + 1].copy()
                    tm.add_task(
                        self.output_text.insert,
                        tk.END,
                        "\n找到表头分隔行，将保留到该行并添加新分析。\n",
                        "info",
                    )
                elif matched_rows:
                    df = df.drop(matched_rows).reset_index(drop=True)
                    tm.add_task(
                        self.output_text.insert,
                        tk.END,
                        f"\n找到 {len(matched_rows)} 行匹配的先前结果，将删除这些行并添加新分析。\n",
                        "warning",
                    )
                elif len(rows_to_keep) < original_count:
                    df = df.loc[rows_to_keep].copy()
                    rows_removed = original_count - len(df)
                    tm.add_task(
                        self.output_text.insert,
                        tk.END,
                        f"\n已删除 {rows_removed} 行匹配的先前结果。\n",
                        "warning",
                    )
                elif original_count >= file_count:
                    # 如果无法确定要删除哪些行，但原始行数大于等于文件数，假设最后几行是要删除的
                    df = df.iloc[:-file_count].copy()

                else:
                    tm.add_task(
                        self.output_text.insert,
                        tk.END,
                        "\n未找到匹配的先前结果，将在Excel末尾添加新分析。建议手动删除可能重复的结果。\n",
                        "info",
                    )

                # 3. 保存更新后的DataFrame
                df.to_excel(self.excel_path.get(), index=False)
                tm.add_task(self.status_bar.configure, text="准备就绪，开始新的分析...")

            # 开始分析
            api_url, api_key = self.get_api_info()
            self.process_papers_async(df, api_url, api_key, file_count)

    except Exception as e:
        error_msg = f"重新生成过程中出错：{str(e)}"
        print(f"重新生成过程异常: {error_msg}")
        import traceback

        print(traceback.format_exc())
        tm.add_task(self.show_error_and_reset, error_msg)


def select_pdfs(self):
    """选择PDF文件并将路径保存到对象中"""
    paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if paths:
        self.pdf_paths = list(paths)  # 转换为列表
        # 清除之前的选择
        if hasattr(self, "selected_pdf_indices"):
            self.selected_pdf_indices = []
        display_pdf_info(self, self.pdf_paths)
        messagebox.showinfo("提示", f"已选择 {len(self.pdf_paths)} 个PDF文件")
        self.status_bar["text"] = f"已选择 {len(self.pdf_paths)} 个PDF文件"
        return True
    return False


def add_pdf(self):
    """添加PDF文件到当前列表"""
    paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if paths:
        if not hasattr(self, "pdf_paths"):
            self.pdf_paths = []

        # 过滤掉已存在的文件
        new_paths = []
        for path in paths:
            if path not in self.pdf_paths:
                new_paths.append(path)
                self.pdf_paths.append(path)

        if new_paths:
            # 清除之前的选择
            if hasattr(self, "selected_pdf_indices"):
                self.selected_pdf_indices = []
            display_pdf_info(self, self.pdf_paths)
            messagebox.showinfo("提示", f"已添加 {len(new_paths)} 个新的PDF文件")
            self.status_bar["text"] = f"当前共选择 {len(self.pdf_paths)} 个PDF文件"
            return True
        else:
            messagebox.showinfo("提示", "所选文件均已存在于列表中")
            return False
    return False


def remove_pdf(self):
    """删除选中的PDF文件"""
    if not hasattr(self, "pdf_paths") or not self.pdf_paths:
        messagebox.showerror("错误", "未选择PDF文件")
        return False

    if not hasattr(self, "selected_pdf_indices") or not self.selected_pdf_indices:
        messagebox.showwarning("警告", "请先点击要删除的PDF文件，选中后再删除")
        return False

    # 确保索引是有效的
    valid_indices = [
        i for i in self.selected_pdf_indices if 0 <= i < len(self.pdf_paths)
    ]

    if not valid_indices:
        messagebox.showwarning("警告", "选择的PDF文件索引无效")
        return False

    # 从大到小排序，避免删除时索引变化
    valid_indices.sort(reverse=True)

    # 删除选定的文件
    deleted_count = 0
    for idx in valid_indices:
        if 0 <= idx < len(self.pdf_paths):
            del self.pdf_paths[idx]
            deleted_count += 1

    # 清空已选择的索引
    self.selected_pdf_indices.clear()

    # 重新显示PDF列表
    display_pdf_info(self, self.pdf_paths)

    # 更新状态栏
    self.status_bar["text"] = (
        f"已删除 {deleted_count} 个文件，当前共有 {len(self.pdf_paths)} 个PDF文件"
    )

    # 如果所有文件都被删除，显示提示信息
    if not self.pdf_paths:
        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, "=== PDF文件列表 ===\n\n", "header")
        self.output_text.insert(tk.END, "未选择任何PDF文件\n", "warning")
        self.output_text.tag_configure(
            "header", foreground="#5ba3e0", font=self.fonts["title"]
        )
        self.output_text.tag_configure("warning", foreground="#e69138")

    return True


def display_pdf_info(self, paths):
    """在界面上显示PDF文件信息"""
    if not hasattr(self, "output_text"):
        return

    self.output_text.delete("1.0", tk.END)
    self.output_text.insert(tk.END, "=== PDF文件列表 ===\n\n", "header")
    self.output_text.tag_configure(
        "header", foreground="#5ba3e0", font=self.fonts["title"]
    )

    if not paths:
        self.output_text.insert(tk.END, "未选择任何PDF文件\n", "warning")
        self.output_text.tag_configure("warning", foreground="#e69138")
        return

    # 删除之前创建的pdf_frame和pdf_listbox（如果有）
    if (
        hasattr(self, "pdf_frame")
        and hasattr(self.pdf_frame, "winfo_exists")
        and self.pdf_frame.winfo_exists()
    ):
        self.pdf_frame.place_forget()
        delattr(self, "pdf_frame")

    if hasattr(self, "pdf_listbox"):
        delattr(self, "pdf_listbox")

    # 设置标签用于跟踪文件位置
    self.pdf_line_positions = {}

    # 显示每个PDF文件及其信息
    for i, path in enumerate(paths, 1):
        filename = os.path.basename(path)
        filesize_kb = os.path.getsize(path) / 1024

        # 保存行号位置信息，用于后续定位文件
        line_position = self.output_text.index(tk.INSERT)
        self.pdf_line_positions[line_position] = i - 1  # 保存对应的索引

        # 为每个文件添加独立的标签
        tag_name = f"pdf_item_{i}"
        self.output_text.insert(
            tk.END, f"{i}. {filename} ({filesize_kb:.1f} KB)\n", (tag_name, "pdf_item")
        )

        # 配置标签能够响应点击事件
        self.output_text.tag_configure(tag_name, foreground="#ffffff")
        self.output_text.tag_bind(
            tag_name, "<Button-1>", lambda e, idx=i - 1: toggle_pdf_selection(self, idx)
        )

    # 创建或清空已选择的PDF文件列表
    if not hasattr(self, "selected_pdf_indices"):
        self.selected_pdf_indices = []
    else:
        self.selected_pdf_indices.clear()

    self.output_text.tag_configure("pdf_item", foreground="#ffffff")
    self.output_text.tag_configure("selected_pdf", background="#3a546d")
    self.output_text.insert(tk.END, f"\n总共 {len(paths)} 个PDF文件\n", "summary")
    self.output_text.tag_configure(
        "summary", foreground="#8bc34a", font=("微软雅黑", 10, "bold")
    )

    # 添加说明文本
    self.output_text.insert(
        tk.END, "\n点击文件名以选择/取消选择，选择后可以删除\n", "instruction"
    )
    self.output_text.tag_configure(
        "instruction", foreground="#aaaaaa", font=("微软雅黑", 9, "italic")
    )


def toggle_pdf_selection(self, index):
    """切换PDF选择状态"""
    if not hasattr(self, "pdf_paths") or not self.pdf_paths:
        return

    if not hasattr(self, "selected_pdf_indices"):
        self.selected_pdf_indices = []

    # 查找对应的标签名称
    tag_name = f"pdf_item_{index+1}"

    # 检查是否已选择
    if index in self.selected_pdf_indices:
        # 取消选择
        self.selected_pdf_indices.remove(index)
        # 移除高亮
        self.output_text.tag_remove("selected_pdf", f"1.0", tk.END)
    else:
        # 添加选择
        self.selected_pdf_indices.append(index)

    # 重新高亮所有选择的项
    self.output_text.tag_remove("selected_pdf", f"1.0", tk.END)
    for idx in self.selected_pdf_indices:
        tag_name = f"pdf_item_{idx+1}"
        # 获取标签范围
        tag_ranges = self.output_text.tag_ranges(tag_name)
        if tag_ranges:
            self.output_text.tag_add("selected_pdf", tag_ranges[0], tag_ranges[1])

    # 更新状态栏信息
    if self.selected_pdf_indices:
        self.status_bar["text"] = f"已选择 {len(self.selected_pdf_indices)} 个PDF文件"
    else:
        self.status_bar["text"] = f"当前共有 {len(self.pdf_paths)} 个PDF文件"
