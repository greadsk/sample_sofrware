import subprocess
import sys
import importlib.machinery
 
path = "/path/to/module"
loader = importlib.machinery.FileFinder(path)
def install_packages():
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade','pyinstaller'])

    # subprocess.check_call([sys.executable, '-m', "ensurepip", "--upgrade"])
    # subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'setuptools'])
    # subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'setuptools==60.0.0'])
    print("pip 更新成功！")
    packages = [
        "numpy",
        "pandas",
        "openpyxl",
        "scipy"
    ]
    for package in packages:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", package,
            "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"
        ])

if __name__ == "__main__":
    install_packages()
    