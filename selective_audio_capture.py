"""
进程音频捕获器 - 仅录制目标进程的音频输出
系统要求: Windows 10 2004+ (build 19041+) 或 Windows 11
"""

import sys
import time
import msvcrt
import threading
import ctypes
from ctypes import wintypes
from pathlib import Path
from datetime import datetime
from typing import Optional, List
from process_audio_capture import ProcessAudioCapture
# pip install process-audio-capture

class PipeAudioSink:
    """命名管道音频接收器，负责接收音频流并实时写入文件，同时动态修复文件头。"""
    
    def __init__(self, output_path: Path):
        self.output_path = output_path
        self._pipe_handle = None
        self._thread = None
        self._stop_event = threading.Event()
        self._pipe_name = fr'\\.\pipe\pac_rec_{int(time.time())}_{id(self)}'
        self._data_size_offset = None
        
    def __enter__(self):
        self.start()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
        
    @property
    def pipe_name(self) -> str:
        return self._pipe_name
        
    def start(self):
        """创建管道并启动接收线程"""
        PIPE_ACCESS_INBOUND = 0x00000001
        PIPE_TYPE_BYTE = 0x00000000
        PIPE_WAIT = 0x00000000
        INVALID_HANDLE_VALUE = -1
        
        self._pipe_handle = ctypes.windll.kernel32.CreateNamedPipeW(
            self._pipe_name,
            PIPE_ACCESS_INBOUND,
            PIPE_TYPE_BYTE | PIPE_WAIT,
            1, 65536, 65536, 0, None
        )
        
        if self._pipe_handle == INVALID_HANDLE_VALUE:
            raise OSError(f"Failed to create named pipe: {ctypes.get_last_error()}")
            
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        
    def stop(self):
        """停止接收并释放资源"""
        self._stop_event.set()
        if self._pipe_handle:
            # 强制断开管道以解除阻塞的读取操作
            ctypes.windll.kernel32.DisconnectNamedPipe(self._pipe_handle)
            ctypes.windll.kernel32.CloseHandle(self._pipe_handle)
            self._pipe_handle = None
        if self._thread:
            self._thread.join(timeout=1.0)

    def _run(self):
        """后台线程：连接管道，读取数据，写入文件，周期性更新头部"""
        ctypes.windll.kernel32.ConnectNamedPipe(self._pipe_handle, None)
        
        buffer = ctypes.create_string_buffer(65536)
        bytes_read = wintypes.DWORD()
        total_written = 0
        last_flush_time = time.time()
        
        try:
            with open(self.output_path, 'wb') as f:
                while not self._stop_event.is_set():
                    success = ctypes.windll.kernel32.ReadFile(
                        self._pipe_handle,
                        buffer,
                        len(buffer),
                        ctypes.byref(bytes_read),
                        None
                    )
                    
                    if not success or bytes_read.value == 0:
                        break
                        
                    data = buffer.raw[:bytes_read.value]
                    
                    # 首次接收数据时解析 data chunk 偏移量
                    if self._data_size_offset is None:
                        self._data_size_offset = self._find_data_chunk_offset(data, total_written)
                            
                    f.write(data)
                    total_written += len(data)
                    
                    # 每秒更新一次文件头，确保文件大小正确
                    now = time.time()
                    if now - last_flush_time > 1.0:
                        f.flush()
                        self._update_wav_header(f, total_written)
                        last_flush_time = now
                
                # 结束时最后更新一次
                f.flush()
                self._update_wav_header(f, total_written)
        except Exception as e:
            sys.stderr.write(f"\n[PipeSink] Error: {e}\n")

    def _find_data_chunk_offset(self, data: bytes, current_offset: int) -> Optional[int]:
        """查找 'data' 块的大小字段偏移量"""
        # 仅在文件头部附近查找
        if current_offset > 200: 
            return None
            
        try:
            data_idx = data.find(b'data')
            if data_idx != -1:
                # 'data' tag (4 bytes) + size (4 bytes) -> size 字段起始位置
                return current_offset + data_idx + 4
        except Exception:
            pass
        return None

    def _update_wav_header(self, f, file_size: int):
        """更新 WAV 文件头中的 RIFF 和 Data chunk 大小"""
        if file_size < 44 or self._data_size_offset is None:
            return
            
        try:
            current_pos = f.tell()
            
            # 1. 更新 RIFF chunk size (File Size - 8)
            f.seek(4)
            f.write((file_size - 8).to_bytes(4, 'little'))
            
            # 2. 更新 Data chunk size
            data_size = file_size - (self._data_size_offset + 4)
            if data_size > 0:
                f.seek(self._data_size_offset)
                f.write(data_size.to_bytes(4, 'little'))
            
            f.seek(current_pos)
        except Exception:
            pass


class AudioRecorder:
    """音频录制管理器"""
    
    def __init__(self, output_dir: str = "./recordings"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self._level_db: float = -60.0
        self._bytes_captured: int = 0
    
    @staticmethod
    def enumerate_processes() -> list:
        return ProcessAudioCapture.enumerate_audio_processes()
    
    def capture(self, pid: int, process_name: str) -> Path:
        """开始捕获流程"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = process_name.rsplit('.', 1)[0] if '.' in process_name else process_name
        output_path = self.output_dir / f"{safe_name}_{timestamp}.wav"
        
        start_time = time.time()
        
        print(f"按 Ctrl+C 停止\n\n")
        
        try:
            # 使用上下文管理器自动处理资源释放
            with PipeAudioSink(output_path) as sink:
                # 启动底层捕获，将音频流导向命名管道
                with ProcessAudioCapture(
                    pid=pid, 
                    output_path=sink.pipe_name, 
                    level_callback=lambda db: setattr(self, '_level_db', db)
                ) as capture:
                    
                    capture.start()
                    # 等待管道建立连接
                    time.sleep(0.2)
                    
                    while True:
                        elapsed = time.time() - start_time
                        self._bytes_captured = output_path.stat().st_size if output_path.exists() else 0
                        self._display_status(elapsed)
                        time.sleep(0.1)
                        
        except KeyboardInterrupt:
            pass
        finally:
            # 恢复光标并换行
            sys.stdout.write("\033[?25h\n\n")
            
        return output_path
    
    def _display_status(self, elapsed: float) -> None:
        """显示实时状态栏"""
        hours, rem = divmod(int(elapsed), 3600)
        mins, secs = divmod(rem, 60)
        
        # 计算音量条
        level_normalized = max(0, min(1, (self._level_db + 60) / 60))
        bars = int(level_normalized * 10)
        level_bar = "█" * bars + "░" * (10 - bars)
        
        # 格式化文件大小
        size = self._bytes_captured
        if size >= 1024**3:
            size_str = f"{size / 1024**3:5.1f}GB"
        else:
            size_str = f"{size / 1024**2:5.1f}MB"
        
        # ANSI 转义：上移一行 + 隐藏光标 + 回车
        status_line = (
            f"\033[1A\033[?25l\r"
            f"[录制] {hours:02d}:{mins:02d}:{secs:02d} [缓存] {size_str}\n"
            f"[分贝] {self._level_db:+6.1f}dB [音量] {level_bar}"
        )
        sys.stdout.write(status_line)
        sys.stdout.flush()


def get_user_selection(options: List) -> Optional[int]:
    """获取用户选择索引 (1-based)"""
    if not options:
        return None
        
    ch = msvcrt.getch()
    if ch == b'\x03': # Ctrl+C
        raise KeyboardInterrupt
    if ch == b'\x1b': # ESC
        return None
        
    try:
        selection = int(ch.decode())
        if 1 <= selection <= len(options):
            return selection
    except (ValueError, UnicodeDecodeError):
        pass
    return None

def format_title(title: str, max_len: int = 24) -> str:
    if not title: return ""
    return title[:max_len - 2] + ".." if len(title) > max_len else title

def main():
    if not ProcessAudioCapture.is_supported():
        sys.stderr.write("错误: 需要 Windows 10 2004+ 或 Windows 11\n")
        return 1
    
    recorder = AudioRecorder()
    
    try:
        processes = recorder.enumerate_processes()
    except Exception as e:
        sys.stderr.write(f"枚举进程失败: {e}\n")
        return 1
    
    if not processes:
        print("未检测到正在播放音频的进程")
        return 0
    
    processes.sort(key=lambda p: p.name.lower())
    
    print("\n音频进程列表:")
    for i, p in enumerate(processes, 1):
        title = format_title(p.window_title or "")
        print(f"  {i}. {p.name:<20} {title}")
    
    print(f"\n选择录制对象: ", end='', flush=True)
    
    selection = get_user_selection(processes)
    if not selection:
        return 0
        
    target = processes[selection - 1]
    
    # 清除选择提示行并显示结果
    # 使用足够多的空格覆盖可能残留的字符
    sys.stdout.write(f"\r选择录制对象: {target.name} [PID:{target.pid}]" + " " * 20 + "\n")
    sys.stdout.flush()
    
    output_path = recorder.capture(target.pid, target.name)
    
    if output_path.exists() and output_path.stat().st_size > 44:
        size_mb = output_path.stat().st_size / (1024 * 1024)
        print(f"录音已保存: {output_path} ({size_mb:.1f}MB)")
    else:
        print("警告: 未捕获到有效音频数据")
        if output_path.exists():
            output_path.unlink()
            
    return 0

if __name__ == "__main__":
    sys.exit(main())
