#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
辅助工具（util_）：PPT 合并/叙事脚本依赖与流程自检。

用于验证 ppt_engine.py 和 merge_ppt.py 的功能是否正常。勿与主程序混淆。
"""

import os
import sys
from pathlib import Path

# 设置输出编码（Windows兼容）
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def test_imports():
    """测试所有必要的导入"""
    print("=" * 80)
    print("测试1: 导入检查")
    print("=" * 80)
    
    try:
        from pptx import Presentation
        print("[OK] python-pptx 导入成功")
    except ImportError as e:
        print(f"[ERROR] python-pptx 导入失败: {e}")
        return False
    
    try:
        import pandas as pd
        print("[OK] pandas 导入成功")
    except ImportError as e:
        print(f"[WARN] pandas 导入失败: {e} (可选)")
    
    try:
        import numpy as np
        print("[OK] numpy 导入成功")
    except ImportError as e:
        print(f"[WARN] numpy 导入失败: {e} (可选)")
    
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
        print("[OK] sklearn 导入成功")
    except ImportError as e:
        print(f"[WARN] sklearn 导入失败: {e} (可选)")
    
    try:
        from merge_ppt import (
            get_merge_dir,
            get_output_dir,
            auto_discover_ppts,
            build_slide_infos,
        )
        print("[OK] merge_ppt 模块导入成功")
    except ImportError as e:
        print(f"[ERROR] merge_ppt 模块导入失败: {e}")
        return False
    
    try:
        from ppt_engine import (
            run_editor,
            run_engine,
            CURATED_SCRIPT,
            SLIDE_BLUEPRINT,
        )
        print("[OK] ppt_engine 模块导入成功")
    except ImportError as e:
        print(f"[ERROR] ppt_engine 模块导入失败: {e}")
        return False
    
    print()
    return True


def test_directory_structure():
    """测试目录结构"""
    print("=" * 80)
    print("测试2: 目录结构检查")
    print("=" * 80)
    
    base_dir = Path(__file__).parent
    input_dir = base_dir / "input"
    output_dir = base_dir / "output"
    
    if input_dir.exists():
        print(f"[OK] 输入目录存在: {input_dir}")
        ppt_files = list(input_dir.glob("*.pptx"))
        print(f"  发现 {len(ppt_files)} 个PPT文件")
        for ppt in ppt_files:
            print(f"    - {ppt.name}")
    else:
        print(f"[WARN] 输入目录不存在: {input_dir}")
        print(f"  创建目录...")
        input_dir.mkdir(exist_ok=True)
    
    if output_dir.exists():
        print(f"[OK] 输出目录存在: {output_dir}")
    else:
        print(f"[WARN] 输出目录不存在: {output_dir}")
        print(f"  创建目录...")
        output_dir.mkdir(exist_ok=True)
    
    print()
    return True


def test_ppt_discovery():
    """测试PPT文件发现功能"""
    print("=" * 80)
    print("测试3: PPT文件发现")
    print("=" * 80)
    
    try:
        from merge_ppt import get_merge_dir, auto_discover_ppts
        
        input_dir = get_merge_dir()
        print(f"输入目录: {input_dir}")
        
        template_path, source_paths = auto_discover_ppts(input_dir)
        print(f"[OK] 模板文件: {os.path.basename(template_path)}")
        print(f"[OK] 源文件数量: {len(source_paths)}")
        
        if len(source_paths) == 0:
            print("[WARN] 警告: 未发现任何源PPT文件")
            return False
        
        return True
    except Exception as e:
        print(f"[ERROR] PPT文件发现失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_slide_extraction():
    """测试幻灯片提取功能"""
    print("=" * 80)
    print("测试4: 幻灯片提取")
    print("=" * 80)
    
    try:
        from merge_ppt import get_merge_dir, auto_discover_ppts, build_slide_infos
        
        input_dir = get_merge_dir()
        template_path, source_paths = auto_discover_ppts(input_dir)
        
        if len(source_paths) == 0:
            print("[WARN] 跳过: 无源文件")
            return True
        
        slide_infos = build_slide_infos(source_paths)
        print(f"[OK] 成功提取 {len(slide_infos)} 张幻灯片")
        
        if len(slide_infos) > 0:
            first = slide_infos[0]
            print(f"  示例: Slide ID {first.global_id} from {first.source_ppt}")
            print(f"  形状数量: {first.shape_count}")
            print(f"  文本长度: {len(first.text_content)} 字符")
        
        return True
    except Exception as e:
        print(f"[ERROR] 幻灯片提取失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主测试流程"""
    print("\n" + "=" * 80)
    print("PPT处理脚本测试与验证")
    print("=" * 80 + "\n")
    
    results = []
    
    # 测试1: 导入
    results.append(("导入检查", test_imports()))
    
    # 测试2: 目录结构
    results.append(("目录结构", test_directory_structure()))
    
    # 测试3: PPT发现
    if results[0][1]:  # 如果导入成功
        results.append(("PPT文件发现", test_ppt_discovery()))
    
    # 测试4: 幻灯片提取
    if results[-1][1] if results else False:  # 如果PPT发现成功
        results.append(("幻灯片提取", test_slide_extraction()))
    
    # 总结
    print("=" * 80)
    print("测试总结")
    print("=" * 80)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"{status}: {name}")
    
    print()
    print(f"总计: {passed}/{total} 项测试通过")
    
    if passed == total:
        print("\n[SUCCESS] 所有测试通过！可以运行主程序。")
        return 0
    else:
        print("\n[WARN] 部分测试未通过，请检查错误信息。")
        return 1


if __name__ == "__main__":
    sys.exit(main())
