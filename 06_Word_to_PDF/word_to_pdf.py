# -*- coding: utf-8 -*-
"""
批量将 Word 文档转换为 PDF。
保留目录书签并支持子文件夹遍历。（已修复 EXE 打包兼容性问题）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List

# 【修复 2】：将局部导入移至全局顶部，确保 PyInstaller 静态分析时能打包该依赖
try:
    import win32com.client  # type: ignore
except ImportError:
    print("错误：需要安装 pywin32。请执行: pip install pywin32", file=sys.stderr)
    sys.exit(1)


def ensure_windows() -> None:
    """确保脚本运行在 Windows 环境下"""
    if sys.platform != "win32":
        print("错误：Word 转 PDF 需要 Windows 系统以及安装 Microsoft Word。", file=sys.stderr)
        sys.exit(1)


# 【修复 1】：增加一个专门用于获取绝对路径的函数，完美兼容 .py 和 .exe 模式
def get_base_dir() -> Path:
    """获取脚本或 EXE 运行所在的真实绝对路径"""
    if getattr(sys, 'frozen', False):
        # 如果是被 PyInstaller 打包成了 EXE 运行
        return Path(sys.executable).parent
    else:
        # 如果是作为普通 .py 脚本运行
        return Path(__file__).resolve().parent


class WordPdfConverter:
    """管理 Word 应用程序生命周期和转换逻辑的类"""
    
    def __init__(self) -> None:
        self._win32 = win32com.client
        self._app = None

    def __enter__(self) -> "WordPdfConverter":
        """进入上下文管理器，启动 Word 进程"""
        self._app = self._win32.Dispatch("Word.Application")
        self._app.Visible = False  # 后台静默运行
        try:
            self._app.DisplayAlerts = 0  # 屏蔽弹窗警告
        except Exception:
            pass
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        """退出上下文管理器，安全关闭 Word 进程"""
        if self._app:
            try:
                self._app.Quit()
            except Exception:
                pass
        self._app = None

    def export(self, doc_path: Path, pdf_path: Path, overwrite: bool = False) -> bool:
        """将单个 Word 文档导出为 PDF"""
        if not doc_path.exists():
            print(f"跳过：输入文件不存在 {doc_path}")
            return False

        if pdf_path.exists() and not overwrite:
            print(f"跳过（已存在）：{pdf_path}")
            return False

        doc = None
        try:
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            doc = self._app.Documents.Open(str(doc_path))
            
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=17,        
                OpenAfterExport=False,  
                OptimizeFor=0,          
                CreateBookmarks=1,      
                DocStructureTags=True   
            )
            
            print(f"已转换：{doc_path} -> {pdf_path}")
            return True
        except Exception as exc:
            print(f"转换失败：{doc_path}，原因：{exc}")
            return False
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass


def collect_docs(input_path: Path, recursive: bool = True) -> List[Path]:
    """收集输入路径下的所有 Word 文档。"""
    allowed = {".doc", ".docx", ".docm"}
    if input_path.is_file() and input_path.suffix.lower() in allowed:
        return [input_path.resolve()]
        
    if input_path.is_dir():
        pattern = "**/*" if recursive else "*"
        files: List[Path] = []
        for p in input_path.glob(pattern):
            if not p.is_file():
                continue
            if p.suffix.lower() not in allowed:
                continue
            if p.name.startswith("~$"):
                continue
            files.append(p.resolve())
        return sorted(files)
    return []


def main() -> None:
    ensure_windows()

    # 【应用修复 1】：使用 get_base_dir() 替代 __file__ 
    base_dir = get_base_dir()
    default_input = base_dir / "input"
    default_output = base_dir / "output"

    parser = argparse.ArgumentParser(description="批量将 Word 文档转换为 PDF")
    parser.add_argument("--input", default=str(default_input), help="输入目录或单个 Word 文件")
    parser.add_argument("--output", default=str(default_output), help="输出目录")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的输出文件")
    parser.add_argument("--recursive", action=argparse.BooleanOptionalAction, default=True, help="是否递归遍历")
    parser.add_argument("--keep-structure", action=argparse.BooleanOptionalAction, default=True, help="保留目录结构")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_dir = Path(args.output).expanduser()

    # 当 EXE 运行时如果发现 input 不存在，自动创建并提醒，避免直接闪退
    if input_path.resolve() == default_input and not input_path.exists():
        input_path.mkdir(parents=True, exist_ok=True)
        print(f"💡 已自动创建输入目录：{input_path}")
        print("请将需要转换的 Word 文档放入该目录后，重新运行本程序。")
        input("按回车键退出...") # 阻止控制台闪退
        return

    doc_files = collect_docs(input_path, recursive=args.recursive)
    if not doc_files:
        print(f"未找到 Word 文件：{input_path}")
        input("按回车键退出...") # 阻止控制台闪退
        return

    total = len(doc_files)
    success = 0
    base_input_dir = input_path.resolve() if input_path.is_dir() else None
    
    try:
        with WordPdfConverter() as converter:
            for doc_path in doc_files:
                if base_input_dir and args.keep_structure:
                    rel = doc_path.relative_to(base_input_dir)
                    pdf_path = (output_dir / rel).with_suffix(".pdf")
                else:
                    pdf_path = output_dir / f"{doc_path.stem}.pdf"
                    
                if converter.export(doc_path, pdf_path, overwrite=args.overwrite):
                    success += 1
    except ImportError:
        return

    print(f"\n✅ 处理完成：{success}/{total} 个文件已成功转换为 PDF！")
    input("按回车键退出...") # 阻止控制台闪退


if __name__ == "__main__":
    main()
