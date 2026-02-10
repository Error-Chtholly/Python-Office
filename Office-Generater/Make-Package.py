import os
import sys
import subprocess
import shutil


def install_pyinstaller():
    """检查并安装 PyInstaller"""
    try:
        import PyInstaller
        print("✅ 检测到 PyInstaller 已安装。")
    except ImportError:
        print("⚠️ 未检测到 PyInstaller，正在尝试自动安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✅ PyInstaller 安装成功！")
        except Exception as e:
            print(f"❌ 安装失败，请手动运行 'pip install pyinstaller'。错误: {e}")
            sys.exit(1)


def build_exe(target_file, icon_path=None, no_console=False):
    """
    执行打包命令
    :param target_file: 目标 py 文件的路径
    :param icon_path: 图标文件 (.ico) 的路径 (可选)
    :param no_console: 是否隐藏控制台窗口 (True为隐藏，适合GUI程序)
    """
    if not os.path.exists(target_file):
        print(f"❌ 错误：找不到文件 '{target_file}'")
        return

    # 获取文件名（不带后缀）
    file_name = os.path.splitext(os.path.basename(target_file))[0]
    output_dir = os.path.join(os.getcwd(), "dist")

    print(f"\n🚀 开始打包: {target_file}")
    print("⏳ 正在分析依赖并生成 EXE，这可能需要几分钟...\n")

    # 构建 PyInstaller 命令
    # -F: 生成单个 EXE 文件
    # --clean: 清理临时文件
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "-F",  # 生成单文件
        "--clean",
        target_file
    ]

    # 是否去除控制台 (黑窗口)
    if no_console:
        cmd.append("--noconsole")  # 或者是 -w
    else:
        cmd.append("--console")

    # 是否添加图标
    if icon_path and os.path.exists(icon_path):
        cmd.extend(["--icon", icon_path])

    # 执行命令
    try:
        # 使用 subprocess 调用命令行
        process = subprocess.run(cmd, text=True)

        if process.returncode == 0:
            exe_path = os.path.join(output_dir, f"{file_name}.exe")
            print("\n" + "=" * 40)
            print(f"✅ 打包成功！")
            print(f"📂 EXE 文件位置: {exe_path}")
            print("=" * 40 + "\n")

            # 清理生成的 .spec 文件和 build 文件夹 (可选)
            cleanup(file_name)
        else:
            print("\n❌ 打包过程中出现错误。")

    except Exception as e:
        print(f"\n❌ 发生异常: {e}")


def cleanup(file_name):
    """清理打包产生的临时文件"""
    try:
        spec_file = f"{file_name}.spec"
        build_folder = "build"
        if os.path.exists(spec_file):
            os.remove(spec_file)
        if os.path.exists(build_folder):
            shutil.rmtree(build_folder)
        print("🧹 已清理临时文件 (spec 和 build 目录)。")
    except Exception:
        pass


if __name__ == "__main__":
    # 1. 检查环境
    install_pyinstaller()

    # 2. 获取用户输入
    print("\n--- Python EXE 打包助手 ---")
    target = input("请输入要打包的 .py 文件路径 (可直接拖入文件): ").strip().replace('"', '')

    # 询问是否需要图标
    use_icon = input("是否指定图标 (.ico)? (输入路径或回车跳过): ").strip().replace('"', '')
    icon = use_icon if use_icon else None

    # 询问是否隐藏控制台
    # 如果你的程序是带界面的(PyQt/Tkinter)，建议选 y；如果是命令行工具，选 n
    console_choice = input("是否隐藏运行时原本的黑窗口 (控制台)? (y/n, 默认n): ").strip().lower()
    hide_console = (console_choice == 'y')

    # 3. 开始打包
    build_exe(target, icon, hide_console)

    input("按回车键退出...")