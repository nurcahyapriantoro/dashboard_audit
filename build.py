import os
import shutil
import subprocess
import sys

def clean_build():
    folders_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['ProjectNavy.spec']
    
    for folder in folders_to_clean:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            print(f"Cleaned: {folder}/")
    
    for file in files_to_clean:
        if os.path.exists(file):
            os.remove(file)
            print(f"Cleaned: {file}")

def build_executable():
    print("=" * 60)
    print("🛡️ Project Navy - Build Script")
    print("=" * 60)
    print()
    
    try:
        import PyInstaller
        print("✅ PyInstaller found")
    except ImportError:
        print("❌ PyInstaller not found. Installing...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    print("\n📁 Cleaning previous build...")
    clean_build()
    
    print("\n🔨 Building executable...")
    
    build_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "DashboardAuditApp",
        "--add-data", f"app.py{os.pathsep}.",
        "--add-data", f"utils.py{os.pathsep}.",
        "--add-data", f"logo_pertamina.png{os.pathsep}.",
        "--hidden-import", "streamlit",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "plotly",
        "--collect-all", "streamlit",
        "--collect-all", "altair",
        "--collect-all", "pydeck",
        "run_app.py"
    ]
    
    result = subprocess.run(build_cmd, capture_output=False)
    
    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("✅ Build successful!")
        print("=" * 60)
        print("\nExecutable location: dist/ProjectNavy.exe")
    else:
        print("\n❌ Build failed!")
        print("Check the error messages above.")

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--clean":
        clean_build()
        print("✅ Clean completed!")
    else:
        build_executable()

if __name__ == "__main__":
    main()
