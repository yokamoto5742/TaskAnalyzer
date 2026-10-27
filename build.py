import subprocess


def build_executable() -> None:

    subprocess.run([
        "pyinstaller",
        "--name=TaskAnalyzer",
        "--windowed",
        "--icon=assets/TaskAnalyzer.ico",
        "--add-data", "utils/config.ini:.",
        "main.py"
    ], check=True)

    print("Executable built successfully.")


if __name__ == "__main__":
    build_executable()
