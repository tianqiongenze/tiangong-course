"""从天工项目提取快照，并检查课程文档中可能过期的数字。

做什么：
1. 读天工仓库的版本、测试数、覆盖率、性能指标、数据规模、文档清单；
2. 写 `upstream/snapshot.json`（机读）与 `upstream/天工项目快照.md`（人读）；
3. 追加 `upstream/CHANGELOG.md`（与上次快照的差异）；
4. 扫描课程文档中出现的旧数字，输出"疑似未更新"清单。

用法：
    python tools/sync_upstream.py --tiangong ../Tiangong-EWA
    python tools/sync_upstream.py --tiangong ../Tiangong-EWA --check-only
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

COURSE = Path(__file__).resolve().parents[1]
UP = COURSE / "upstream"

# 需要扫描"旧数字"的课程文档（相对课程根目录）
SCAN_FILES = [
    "README.md",
    "docs/00_课程总览与学习路径.md",
    "docs/01_课程标准.md",
    "docs/06_评价量规与考核题库.md",
    "docs/07_实验环境与工具链.md",
    "docs/03_教学方案/S10_性能工程预聚合与多级缓存.md",
    "docs/03_教学方案/S12_测试工程TDD_BDD_E2E.md",
    "docs/09_学生工作页.md",
    "docs/13_与天工项目同步机制.md",
]


def _git(repo: Path, *args: str) -> str:
    try:
        out = subprocess.run(["git", "-C", str(repo), *args], capture_output=True,
                             text=True, timeout=20, check=False)
        return out.stdout.strip()
    except (OSError, subprocess.SubprocessError):
        return ""


def _read(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except OSError:
        return ""


def collect(tiangong: Path) -> dict:
    """从上游仓库提取可量化事实（失败则留空，不编造）。"""
    snap: dict = {
        "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "repo": "tianqiongenze/tiangong-ewa",
        "path": str(tiangong),
        "commit": _git(tiangong, "rev-parse", "--short", "HEAD"),
        "commit_date": _git(tiangong, "log", "-1", "--format=%cd",
                            "--date=format:%Y-%m-%d"),
        "version": "",
        "tests": None,
        "e2e": None,
        "coverage": None,
        "data_rows": None,
        "data_size": None,
        "perf_7d": None,
        "docs": [],
    }

    # 版本：从 impl/services/agent-runtime/app/__init__.py 读 __version__
    init = _read(tiangong / "impl/services/agent-runtime/app/__init__.py")
    m = re.search(r'__version__\s*=\s*"([^"]+)"', init)
    if m:
        snap["version"] = m.group(1)

    # 测试数 / 覆盖率：从 README 读（人写的地方通常最新）
    readme = _read(tiangong / "README.md")
    m = re.search(r"(\d+)\s*项单元/接口", readme)
    if m:
        snap["tests"] = int(m.group(1))
    m = re.search(r"(\d+)\s*项浏览器\s*E2E", readme)
    if m:
        snap["e2e"] = int(m.group(1))
    m = re.search(r"文档基线：\*\*(v[\d.]+)", readme)
    if m:
        snap["version"] = m.group(1)
    m = re.search(r"覆盖率\s*\*{0,2}(\d+)%", readme)
    if m:
        snap["coverage"] = int(m.group(1))

    # 数据规模与性能：从 docs/44 读
    d44 = _read(tiangong / "docs/44_全链路与性能测试报告.md")
    m = re.search(r"\*\*(\d[\d,]*)\s*行\s*/\s*([\d.]+)\s*GiB\*\*", d44)
    if m:
        snap["data_rows"] = int(m.group(1).replace(",", ""))
        snap["data_size"] = float(m.group(2))
    m = re.search(r"近 7 天概览 \| ([\d,]+) ms \| \*\*([\d.]+) ms\*\*", d44)
    if m:
        snap["perf_7d"] = {"raw_ms": int(m.group(1).replace(",", "")),
                           "rollup_ms": float(m.group(2))}

    # 文档清单
    docs_dir = tiangong / "docs"
    if docs_dir.is_dir():
        snap["docs"] = sorted(p.name for p in docs_dir.glob("*.md"))

    return snap


def render_md(snap: dict) -> str:
    perf = snap.get("perf_7d") or {}
    rows = snap.get("data_rows")
    lines = [
        "# 天工项目快照",
        "",
        f"> 采集时间：{snap['collected_at']}　｜　采集方式：`tools/sync_upstream.py`",
        "> 本文件由脚本生成，**请勿手工编辑**；需要改内容请改上游或改脚本。",
        "",
        "| 项 | 值 |",
        "|---|---|",
        f"| 仓库 | {snap['repo']} |",
        f"| 版本 | {snap.get('version') or '—'} |",
        f"| commit | `{snap.get('commit') or '—'}`（{snap.get('commit_date') or '—'}） |",
        f"| 单元/接口测试 | {snap.get('tests') if snap.get('tests') is not None else '—'} 项 |",
        f"| 浏览器 E2E | {snap.get('e2e') if snap.get('e2e') is not None else '—'} 项 |",
        f"| 覆盖率 | {str(snap.get('coverage')) + '%' if snap.get('coverage') else '—'} |",
        f"| 数据规模 | {f'{rows:,} 行 / {snap.get('data_size')} GiB' if rows else '—'} |",
        f"| 7 天聚合（明细 → 预聚合） | "
        f"{f'{perf.get('raw_ms'):,} ms → {perf.get('rollup_ms')} ms' if perf else '—'} |",
        "",
        "## 课程中引用的上游文档",
        "",
    ]
    for d in snap.get("docs", []):
        lines.append(f"- `docs/{d}`")
    lines.append("")
    return "\n".join(lines)


def diff(prev: dict | None, cur: dict) -> list[str]:
    if not prev:
        return ["首次采集"]
    out = []
    for k in ("version", "commit", "tests", "e2e", "coverage", "data_rows",
              "data_size", "perf_7d"):
        a, b = prev.get(k), cur.get(k)
        if a != b:
            out.append(f"- `{k}`: {a} → {b}")
    return out or ["无变化"]


def scan_stale(cur: dict, prev: dict | None) -> list[str]:
    """在课程文档中找仍写着旧数字的地方。"""
    if not prev:
        return []
    warns: list[str] = []
    patterns = []
    if prev.get("tests") and cur.get("tests") and prev["tests"] != cur["tests"]:
        patterns.append((f"{prev['tests']} 项", f"测试数应为 {cur['tests']} 项"))
    if prev.get("coverage") and cur.get("coverage") and prev["coverage"] != cur["coverage"]:
        patterns.append((f"{prev['coverage']}%", f"覆盖率应为 {cur['coverage']}%"))
    if prev.get("data_rows") and cur.get("data_rows") and prev["data_rows"] != cur["data_rows"]:
        patterns.append((f"{prev['data_rows']:,}", f"行数应为 {cur['data_rows']:,}"))

    for rel in SCAN_FILES:
        p = COURSE / rel
        text = _read(p)
        if not text:
            continue
        for old, hint in patterns:
            if old in text:
                warns.append(f"{rel}: 仍出现「{old}」→ {hint}")
    return warns


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tiangong", required=True, help="天工仓库本地路径")
    ap.add_argument("--check-only", action="store_true", help="只检查，不写文件")
    args = ap.parse_args()

    repo = Path(args.tiangong).expanduser().resolve()
    if not (repo / "README.md").exists():
        print(f"不是天工仓库: {repo}", file=sys.stderr)
        return 2

    UP.mkdir(parents=True, exist_ok=True)
    snap_file = UP / "snapshot.json"
    prev = json.loads(_read(snap_file) or "null")
    cur = collect(repo)

    print("== 上游快照 ==")
    for k in ("version", "commit", "tests", "e2e", "coverage",
              "data_rows", "data_size", "perf_7d"):
        print(f"  {k:<12} {cur.get(k)}")

    changes = diff(prev, cur)
    print("\n== 与上次相比 ==")
    for c in changes:
        print(f"  {c}")

    warns = scan_stale(cur, prev)
    if warns:
        print("\n== 疑似未更新（请逐项核对）==")
        for w in warns:
            print(f"  ! {w}")
    else:
        print("\n== 疑似未更新 ==\n  无")

    if args.check_only:
        return 0

    snap_file.write_text(json.dumps(cur, ensure_ascii=False, indent=2),
                         encoding="utf-8")
    (UP / "天工项目快照.md").write_text(render_md(cur), encoding="utf-8")

    log = UP / "CHANGELOG.md"
    header = "# 上游同步记录\n\n" if not log.exists() else ""
    entry = (f"## {datetime.now().strftime('%Y-%m-%d %H:%M')} · "
             f"{cur.get('version') or '未知版本'} (`{cur.get('commit')}`)\n\n"
             + "\n".join(changes) + "\n\n")
    with log.open("a", encoding="utf-8") as f:
        f.write(header + entry)

    print(f"\n已写入: {snap_file.name} / 天工项目快照.md / CHANGELOG.md")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
