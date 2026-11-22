# -*- coding: utf-8 -*-
"""为冻结目录生成相对路径 SHA-256 清单，便于发布完整性校验。"""
import hashlib
import os
import sys


def write_checksums(root: str) -> str:
    root = os.path.abspath(root)
    rows = []
    for folder, _dirs, files in os.walk(root):
        for name in files:
            if name == "SHA256SUMS.txt":
                continue
            path = os.path.join(folder, name)
            digest = hashlib.sha256()
            with open(path, "rb") as handle:
                for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                    digest.update(chunk)
            relative = os.path.relpath(path, root).replace(os.sep, "/")
            rows.append(f"{digest.hexdigest()} *{relative}")
    target = os.path.join(root, "SHA256SUMS.txt")
    with open(target, "w", encoding="utf-8", newline="\n") as handle:
        handle.write("\n".join(sorted(rows)) + "\n")
    return target


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("用法：python tools/write_checksums.py dist\\PDFDuplexPrint")
    print(write_checksums(sys.argv[1]))
