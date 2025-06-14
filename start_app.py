import os
import sys
import time
import psutil
import subprocess
import webbrowser
import socket
from pathlib import Path

def is_port_in_use(port):
    """检查端口是否被占用"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(('127.0.0.1', port))
            return False
        except socket.error:
            return True

def kill_process_on_port(port):
    """终止占用指定端口的进程"""
    killed = False
    for proc in psutil.process_iter(['pid', 'name', 'connections']):
        try:
            for conn in proc.connections():
                if conn.laddr.port == port:
                    print(f"正在终止占用端口 {port} 的进程 (PID: {proc.pid}, 名称: {proc.name()})")
                    try:
                        psutil.Process(proc.pid).terminate()
                        killed = True
                    except psutil.NoSuchProcess:
                        pass
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    
    if killed:
        time.sleep(2)  # 等待进程完全终止
        return True
    return False

def find_available_port(start_port=8501, max_attempts=10):
    """查找可用端口"""
    port = start_port
    for _ in range(max_attempts):
        if is_port_in_use(port):
            if kill_process_on_port(port):
                print(f"已释放端口 {port}")
            else:
                print(f"端口 {port} 仍被占用，尝试下一个端口")
                port += 1
                continue
        
        try:
            # 再次检查端口是否可用
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                print(f"端口 {port} 可用")
                return port
        except socket.error:
            port += 1
    
    raise RuntimeError(f"无法找到可用端口 (尝试范围: {start_port}-{start_port+max_attempts-1})")

def check_network():
    """检查网络连接"""
    try:
        # 尝试连接本地回环地址
        socket.create_connection(("127.0.0.1", 0), timeout=1)
        return True
    except OSError:
        return False

def start_streamlit():
    """启动Streamlit应用"""
    try:
        # 检查网络连接
        if not check_network():
            print("警告：无法建立本地网络连接，请检查网络设置")
            input("按回车键继续...")
        
        # 获取当前脚本所在目录
        current_dir = Path(__file__).parent.absolute()
        
        # 查找可用端口
        try:
            port = find_available_port()
            print(f"使用端口: {port}")
        except RuntimeError as e:
            print(f"错误: {str(e)}")
            print("请尝试手动关闭占用端口的程序后重试")
            input("按回车键退出...")
            return
        
        # 构建启动命令
        cmd = [
            sys.executable,  # Python解释器路径
            "-m", "streamlit", "run",
            str(current_dir / "app.py"),
            "--server.port", str(port),
            "--server.address", "127.0.0.1",  # 使用IP地址而不是localhost
            "--browser.serverAddress", "127.0.0.1",
            "--server.headless", "true",
            "--server.enableCORS", "false",
            "--server.enableXsrfProtection", "false",
            "--server.maxUploadSize", "10",
            "--server.maxMessageSize", "200"
        ]
        
        print("正在启动应用...")
        print(f"启动命令: {' '.join(cmd)}")
        
        # 启动Streamlit进程
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NEW_CONSOLE  # 在新窗口中运行
        )
        
        # 等待应用启动
        print("等待应用启动...")
        time.sleep(5)  # 增加等待时间
        
        # 检查进程是否还在运行
        if process.poll() is None:
            # 尝试连接应用
            max_retries = 3
            for i in range(max_retries):
                try:
                    with socket.create_connection(("127.0.0.1", port), timeout=2):
                        print("应用已成功启动！")
                        break
                except (socket.timeout, socket.error):
                    if i < max_retries - 1:
                        print(f"等待应用响应... ({i+1}/{max_retries})")
                        time.sleep(2)
                    else:
                        print("警告：应用可能未正确启动，但仍在尝试打开浏览器")
            
            # 打开浏览器
            url = f"http://127.0.0.1:{port}"
            print(f"正在打开浏览器: {url}")
            webbrowser.open(url)
            
            # 显示进程输出
            while True:
                output = process.stdout.readline()
                if output:
                    print(output.strip())
                if process.poll() is not None:
                    break
            
            # 如果进程意外结束，显示错误信息
            if process.poll() != 0:
                error = process.stderr.read()
                print(f"应用异常退出: {error}")
        else:
            error = process.stderr.read()
            print(f"应用启动失败: {error}")
            
    except Exception as e:
        print(f"启动过程中出现错误: {str(e)}")
        print("详细错误信息:", sys.exc_info())
        input("按回车键退出...")
        sys.exit(1)

if __name__ == "__main__":
    print("="*50)
    print("PPT制作工具启动器")
    print("="*50)
    
    # 检查必要的包是否已安装
    required_packages = ['streamlit', 'psutil']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"正在安装缺失的包: {', '.join(missing_packages)}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
        except subprocess.CalledProcessError as e:
            print(f"安装包失败: {str(e)}")
            input("按回车键退出...")
            sys.exit(1)
    
    start_streamlit() 