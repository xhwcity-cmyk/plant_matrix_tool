# build.py
import platform
import subprocess
import sys
import os
import shutil
from pathlib import Path

# 全局常量定义
APP_NAME = "SpeciesProcessor"
FINAL_APP_NAME = "物种数据整理工具"
MAIN_SCRIPT = "plant_matrix_tool.py"


def check_requirements():
    """检查必要的依赖"""
    # 主程序运行需要的包
    required_packages = ['pandas', 'openpyxl', 'numpy']
    missing_packages = []

    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print("❌ 缺少必要的依赖包:")
        for package in missing_packages:
            print(f"  - {package}")
        print(f"\n请运行: pip install {' '.join(missing_packages)}")
        return False

    # 检查主脚本是否存在
    if not os.path.exists(MAIN_SCRIPT):
        print(f"❌ 找不到主脚本文件: {MAIN_SCRIPT}")
        return False

    # 检查 PyInstaller 是否可用（但不作为必要条件）
    try:
        import PyInstaller
        print("✅ PyInstaller 可用")
    except ImportError:
        print("⚠️ PyInstaller 未安装，尝试使用命令行调用")

    return True


def check_pyinstaller_available():
    """检查 PyInstaller 是否在 PATH 中可用"""
    try:
        # 尝试运行 pyinstaller --version
        result = subprocess.run([sys.executable, "-m", "PyInstaller", "--version"],
                                capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print("✅ PyInstaller 命令行可用")
            return True
    except (subprocess.TimeoutExpired, subprocess.SubprocessError, FileNotFoundError):
        pass

    # 尝试直接调用 pyinstaller
    try:
        result = subprocess.run(["pyinstaller", "--version"],
                                capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print("✅ PyInstaller 命令行可用")
            return True
    except (subprocess.TimeoutExpired, subprocess.SubprocessError, FileNotFoundError):
        pass

    print("❌ 无法找到 PyInstaller，请尝试以下方法:")
    print("   1. 使用完整路径: python -m PyInstaller ...")
    print("   2. 重新安装: pip install --upgrade pyinstaller")
    print("   3. 检查 Python 环境")
    return False


def build_app():
    """构建应用程序"""
    if not check_requirements():
        return False

    if not check_pyinstaller_available():
        return False

    current_os = platform.system()
    print(f"检测到操作系统: {current_os}")

    # 尝试不同的 PyInstaller 调用方式
    pyinstaller_commands = [
        [sys.executable, "-m", "PyInstaller"],  # 方式1: 使用模块调用
        ["pyinstaller"]  # 方式2: 直接调用
    ]

    # 基本参数
    base_args = [
        "--onefile",
        "--windowed",
        "--name", APP_NAME,
        "--distpath", "dist",
        "--workpath", "build",
        "--specpath", ".",
        "--clean",
        "--noconfirm"
    ]

    # 添加图标
    icon_path = get_icon_path(current_os)
    if icon_path:
        base_args.extend(["--icon", str(icon_path)])
        print(f"使用图标: {icon_path}")

    # 添加隐藏导入
    hidden_imports = [
        'pandas', 'openpyxl', 'numpy', 'tkinter',
        'tkinter.filedialog', 'tkinter.messagebox',
        'pkg_resources.py2_warn'
    ]

    for imp in hidden_imports:
        base_args.extend(["--hidden-import", imp])

    # 添加数据文件
    data_files = get_data_files()
    for src, dest in data_files:
        sep = ";" if current_os == "Windows" else ":"
        base_args.extend(["--add-data", f"{src}{sep}{dest}"])

    # 添加主脚本
    base_args.append(MAIN_SCRIPT)

    # 尝试不同的命令
    success = False
    for pyinstaller_cmd in pyinstaller_commands:
        cmd = pyinstaller_cmd + base_args

        print(f"\n尝试命令: {' '.join(cmd)}")
        print("开始打包，这可能需要几分钟...")

        try:
            result = subprocess.run(cmd, check=True, capture_output=True, text=True, timeout=300)
            if result.returncode == 0:
                print("✅ 打包成功！")
                success = True
                break
            else:
                print(f"❌ 打包失败，返回码: {result.returncode}")
                if result.stderr:
                    print("错误信息:")
                    print(result.stderr)

        except subprocess.TimeoutExpired:
            print("❌ 打包超时，请重试")
        except subprocess.CalledProcessError as e:
            print(f"❌ 打包失败: {e}")
            if e.stdout:
                print("标准输出:")
                print(e.stdout)
            if e.stderr:
                print("错误输出:")
                print(e.stderr)
        except FileNotFoundError:
            print(f"❌ 命令未找到: {pyinstaller_cmd[0]}")
        except Exception as e:
            print(f"❌ 打包过程中发生错误: {str(e)}")
            import traceback
            traceback.print_exc()

    if success:
        return rename_final_app(current_os)
    else:
        print("\n所有打包尝试都失败了，请检查:")
        print("1. PyInstaller 是否正确安装")
        print("2. 依赖包是否完整")
        print("3. 系统权限是否足够")
        return False


def get_icon_path(current_os):
    """获取图标路径"""
    icon_files = [
        "app_icon.ico",
        "icon.ico",
        "plant.ico",
        "app_icon.png"
    ]

    for icon_file in icon_files:
        if os.path.exists(icon_file):
            return icon_file

    print("⚠️ 未找到图标文件，将使用默认图标")
    return None


def get_data_files():
    """获取需要包含的数据文件"""
    data_files = []

    # 尝试添加 openpyxl 数据文件
    try:
        import openpyxl
        openpyxl_path = os.path.dirname(openpyxl.__file__)
        data_files.append((openpyxl_path, "openpyxl"))
        print(f"添加 openpyxl 数据文件: {openpyxl_path}")
    except ImportError:
        print("⚠️ 未找到 openpyxl，跳过数据文件")

    return data_files


def rename_final_app(current_os):
    """重命名最终应用程序"""
    try:
        if current_os == "Darwin":
            original_path = Path("dist") / f"{APP_NAME}.app"
            final_path = Path("dist") / f"{FINAL_APP_NAME}.app"

            if original_path.exists():
                if final_path.exists():
                    shutil.rmtree(final_path)
                original_path.rename(final_path)
                print(f"✅ 重命名为: {final_path}")
                return True
            else:
                print("⚠️ 警告: 未找到生成的 .app 文件")
                return False

        elif current_os == "Windows":
            original_path = Path("dist") / f"{APP_NAME}.exe"
            final_path = Path("dist") / f"{FINAL_APP_NAME}.exe"

            if original_path.exists():
                if final_path.exists():
                    os.remove(final_path)
                original_path.rename(final_path)
                print(f"✅ 重命名为: {final_path}")
                return True
            else:
                print("⚠️ 警告: 未找到生成的 .exe 文件")
                return False

        else:
            original_path = Path("dist") / APP_NAME
            final_path = Path("dist") / FINAL_APP_NAME

            if original_path.exists():
                if final_path.exists():
                    os.remove(final_path)
                original_path.rename(final_path)
                print(f"✅ 重命名为: {final_path}")
                return True
            else:
                print("⚠️ 警告: 未找到生成的可执行文件")
                return False

    except Exception as e:
        print(f"❌ 重命名失败: {str(e)}")
        return False


def clean_build():
    """清理构建目录"""
    print("清理旧构建文件...")

    cleanup_items = [
        "build",
        "dist",
        f"{APP_NAME}.spec",
        "__pycache__"
    ]

    for item in cleanup_items:
        if os.path.exists(item):
            try:
                if os.path.isdir(item):
                    shutil.rmtree(item)
                    print(f"✅ 已清理目录: {item}")
                else:
                    os.remove(item)
                    print(f"✅ 已清理文件: {item}")
            except Exception as e:
                print(f"⚠️ 清理 {item} 失败: {str(e)}")


def show_final_instructions(current_os):
    """显示最终使用说明"""
    print("\n" + "=" * 50)
    print("🎉 打包完成！")
    print("=" * 50)

    if current_os == "Darwin":
        app_path = Path("dist") / f"{FINAL_APP_NAME}.app"
        if app_path.exists():
            print(f"macOS 用户:")
            print(f"  1. 将 '{app_path.name}' 拖到'应用程序'文件夹")
            print(f"  2. 在Launchpad或应用程序文件夹中打开")
        else:
            print("⚠️ 警告: 未找到生成的应用程序文件")

    elif current_os == "Windows":
        exe_path = Path("dist") / f"{FINAL_APP_NAME}.exe"
        if exe_path.exists():
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"Windows 用户:")
            print(f"  1. 将 '{exe_path.name}' 发送给用户")
            print(f"  2. 用户双击即可运行，无需安装Python")
            print(f"  文件大小: {size_mb:.1f} MB")
        else:
            print("⚠️ 警告: 未找到生成的应用程序文件")

    else:
        app_path = Path("dist") / FINAL_APP_NAME
        if app_path.exists():
            print(f"Linux/其他系统用户:")
            print(f"  可执行文件: {app_path}")
            print(f"  可能需要执行: chmod +x {app_path}")

    print(f"\n输出目录: {Path('dist').absolute()}")
    print("=" * 50)


if __name__ == "__main__":
    print("开始打包应用程序...")
    print(f"应用程序名称: {FINAL_APP_NAME}")
    print(f"主脚本: {MAIN_SCRIPT}")

    # 获取当前操作系统
    current_os = platform.system()

    # 清理旧构建
    clean_build()

    # 开始打包
    success = build_app()

    if success:
        show_final_instructions(current_os)
    else:
        print("\n❌ 打包失败，请检查上面的错误信息")
        sys.exit(1)