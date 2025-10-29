import zipfile
from pathlib import Path  # 用于路径处理（可选，也可用字符串路径）

def zip_files(file_paths, zip_name):
    """
    打包多个文件为 ZIP 压缩包
    :param file_paths: 要打包的文件路径列表（支持字符串或 Path 对象）
    :param zip_name: 输出的 ZIP 文件名（如 "output.zip"）
    """
    # 创建 ZIP 文件（mode='w' 表示写入，zipfile.ZIP_DEFLATED 表示启用压缩）
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in file_paths:
            # 转换为 Path 对象便于处理
            path = Path(file_path)
            if not path.exists():
                print(f"警告：文件不存在，跳过打包：{path}")
                continue
            # 将文件添加到 ZIP 中，arcname 可指定在 ZIP 中的相对路径（默认保留原路径结构）
            zipf.write(path, arcname=path.name)  # 仅保留文件名，不包含上级目录
            print(f"已添加：{path}")
    print(f"打包完成：{zip_name}")

# 测试：打包多个文件
if __name__ == "__main__":
    files_to_zip = [
        "file1.txt",
        "data/report.csv",
        "images/photo.jpg"
    ]
    zip_files(files_to_zip, "output.zip")