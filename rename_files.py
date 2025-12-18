#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
MDファイル名にCSVから取得した案件番号を付与してoutputディレクトリにコピーする

ファイル名の形式: PJ名_フェーズ名_成果物種別_... → 【案件番号】PJ名_フェーズ名_...
CSVの案件名とPJ名を部分一致で照合し、一致しない場合は【不明】を付与する
"""

import csv
import os
import shutil
from pathlib import Path


def load_project_mapping(csv_path: str) -> dict[str, str]:
    """
    CSVから案件名 → 案件番号のマッピングを作成する
    同じ案件名に対して複数の案件番号がある場合は、最初に見つかったものを使用
    """
    mapping = {}
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader)  # ヘッダ行をスキップ
        
        for row in reader:
            if len(row) >= 3:
                project_number = row[0].strip()
                project_name = row[2].strip()
                
                # 案件番号が空でなく、案件名がまだ登録されていない場合のみ追加
                if project_number and project_name and project_name not in mapping:
                    mapping[project_name] = project_number
    
    return mapping


def normalize_text(text: str) -> str:
    """
    テキストを正規化する（全角・半角統一、括弧除去など）
    """
    import unicodedata
    # 全角→半角
    normalized = unicodedata.normalize("NFKC", text)
    # 括弧とその中身を除去するパターン
    import re
    normalized = re.sub(r"[【】\[\]（）\(\)]", "", normalized)
    return normalized.lower()


def extract_keywords(text: str) -> set[str]:
    """
    テキストからキーワードを抽出する
    """
    normalized = normalize_text(text)
    # 一般的な区切り文字で分割
    import re
    keywords = set(re.split(r"[_\s\-・]+", normalized))
    # 空文字を除去
    keywords.discard("")
    return keywords


def find_project_number(pj_name: str, mapping: dict[str, str]) -> str:
    """
    PJ名からCSVの案件名を検索し、一致する案件番号を返す
    完全一致 → 部分一致 → キーワードマッチの順で検索
    一致しない場合は "不明" を返す
    """
    # 完全一致を探す
    for project_name, project_number in mapping.items():
        if pj_name == project_name:
            return project_number
    
    # 正規化して比較
    normalized_pj = normalize_text(pj_name)
    for project_name, project_number in mapping.items():
        normalized_csv = normalize_text(project_name)
        if normalized_pj == normalized_csv:
            return project_number
    
    # 部分一致を探す（CSVの案件名にPJ名が含まれる場合）
    for project_name, project_number in mapping.items():
        if pj_name in project_name:
            return project_number
    
    # 逆に、PJ名にCSVの案件名が含まれる場合
    for project_name, project_number in mapping.items():
        if project_name in pj_name:
            return project_number
    
    # 正規化後の部分一致
    for project_name, project_number in mapping.items():
        normalized_csv = normalize_text(project_name)
        if normalized_pj in normalized_csv or normalized_csv in normalized_pj:
            return project_number
    
    # キーワードマッチ（PJ名のキーワードがすべてCSVの案件名に含まれる場合）
    pj_keywords = extract_keywords(pj_name)
    if pj_keywords:
        for project_name, project_number in mapping.items():
            csv_keywords = extract_keywords(project_name)
            # PJ名の主要キーワードがCSV案件名に含まれているか
            if pj_keywords.issubset(csv_keywords):
                return project_number
    
    return "不明"


def extract_pj_name(filename: str) -> str:
    """
    ファイル名からPJ名を抽出する
    形式: PJ名_フェーズ名_... → 最初のアンダースコアまで
    """
    parts = filename.split("_")
    if parts:
        return parts[0]
    return filename


def truncate_filename(filename: str, max_length: int = 200) -> str:
    """
    ファイル名が長すぎる場合に短縮する
    ファイル名の一意性を保つためにハッシュを付加
    """
    import hashlib
    
    if len(filename.encode('utf-8')) <= max_length:
        return filename
    
    # 拡張子を取り出す
    base, ext = os.path.splitext(filename)
    
    # ハッシュを生成（元のファイル名の一意性を保つため）
    file_hash = hashlib.md5(filename.encode('utf-8')).hexdigest()[:8]
    
    # 利用可能な長さを計算（ハッシュ + 拡張子 + アンダースコアの分を引く）
    available_length = max_length - len(ext.encode('utf-8')) - len(file_hash) - 1
    
    # ベース名を切り詰める（UTF-8バイト数で計算）
    truncated_base = base
    while len(truncated_base.encode('utf-8')) > available_length:
        truncated_base = truncated_base[:-1]
    
    return f"{truncated_base}_{file_hash}{ext}"


def process_files(input_dir: str, output_dir: str, csv_path: str):
    """
    入力ディレクトリのMDファイルを処理し、案件番号を付与して出力ディレクトリにコピーする
    """
    # 出力ディレクトリを作成
    os.makedirs(output_dir, exist_ok=True)
    
    # CSVからマッピングを読み込む
    mapping = load_project_mapping(csv_path)
    print(f"CSVから {len(mapping)} 件の案件を読み込んだ")
    
    # 入力ディレクトリのMD/PDFファイルを処理
    input_path = Path(input_dir)
    md_files = list(input_path.glob("*.md")) + list(input_path.glob("*.pdf"))
    print(f"処理対象: {len(md_files)} 件のファイル（MD/PDF）")
    
    # 統計情報
    matched = 0
    unmatched = 0
    truncated = 0
    
    for md_file in md_files:
        filename = md_file.name
        pj_name = extract_pj_name(filename)
        project_number = find_project_number(pj_name, mapping)
        
        # 新しいファイル名を作成
        new_filename = f"【{project_number}】{filename}"
        
        # ファイル名が長すぎる場合は短縮
        original_new_filename = new_filename
        new_filename = truncate_filename(new_filename)
        if new_filename != original_new_filename:
            truncated += 1
        
        output_path = Path(output_dir) / new_filename
        
        # ファイルをコピー
        shutil.copy2(md_file, output_path)
        
        if project_number == "不明":
            unmatched += 1
            print(f"[不明] {filename} (PJ名: {pj_name})")
        else:
            matched += 1
            print(f"[{project_number}] {filename}")
    
    print(f"\n処理完了:")
    print(f"  一致: {matched} 件")
    print(f"  不明: {unmatched} 件")
    print(f"  ファイル名短縮: {truncated} 件")


def main():
    # パスの設定
    base_dir = Path(__file__).parent
    input_dir = base_dir.parent / "10_PJKB取込用資料(名前補正済み)" / "まとめ"
    output_dir = base_dir / "data" / "output"
    csv_path = base_dir / "data" / "input" / "クライテリア管理台帳.csv"
    
    print(f"入力ディレクトリ: {input_dir}")
    print(f"出力ディレクトリ: {output_dir}")
    print(f"CSVファイル: {csv_path}")
    print()
    
    process_files(str(input_dir), str(output_dir), str(csv_path))


if __name__ == "__main__":
    main()
