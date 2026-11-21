"""
버전별 패키지 관리

    python app/release.py            지금 버전으로 패키지 만들기  → releases/SellerMorning_v1.2.0_2026-11-06.zip
    python app/release.py list       만든 패키지 목록
    python app/release.py restore 1.0.0   그 버전으로 되돌리기(지금 파일은 먼저 releases/backup/ 에 묶어 둠)

- 들어가는 것: app/ 의 프로그램·자료 파일(.py · .json · assets), 처음 폴더의 실행 파일(.bat · .command)·README·requirements·.gitignore
- 넣지 않는 것: reports/(판매 데이터), __pycache__, .venv, 정리용 cleanup_*.bat
  로그인 정보·API 키·Seller ID 등은 원래 프로젝트 밖(%LOCALAPPDATA%\\AmazonAIProject)에 있어서 들어가지 않습니다.
- 만들기 전에 검사: 설정에 저장된 Seller ID·메일·브랜드·비밀값이 파일 안에 들어 있으면 만들지 않음
- 패키지 안의 RELEASE.txt: 버전, 날짜, 파일마다 크기·SHA-256(나중에 바뀐 파일 확인용)
- 같은 버전 패키지가 이미 있고 내용이 다르면 만들지 않음 → app/version.py에서 버전을 올리라고 안내
"""

import hashlib
import io
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
OUT = ROOT / "releases"
NAME = "SellerMorning"
OLD_NAMES = ("AmazonAIProject",)   # 이름을 바꾸기 전에 만든 패키지도 목록·되돌리기에 보이게

APP_EXT = {".py", ".json", ".ico", ".png"}
ROOT_FILES = {"README.md", "requirements.txt", ".gitignore"}
ROOT_EXT = {".bat", ".command"}
SKIP_DIRS = {"__pycache__", ".venv", "venv", "reports", "releases", "browser_profile", ".git"}


def files() -> List[Path]:
    out: List[Path] = []
    for p in sorted(BASE.rglob("*")):
        rel = p.relative_to(ROOT)
        if p.is_file() and p.suffix.lower() in APP_EXT and not (set(rel.parts) & SKIP_DIRS):
            out.append(p)
    for p in sorted(ROOT.iterdir()):
        if p.is_file() and (p.name in ROOT_FILES or p.suffix.lower() in ROOT_EXT) and not p.name.lower().startswith("cleanup"):
            out.append(p)
    return out


def _glob(rest: str) -> List[Path]:
    """지금 이름과 예전 이름(OLD_NAMES)의 패키지를 함께 찾음."""
    if not OUT.exists():
        return []
    out: List[Path] = []
    for n in (NAME,) + OLD_NAMES:
        out += OUT.glob(f"{n}_v{rest}")
    return sorted(set(out), key=lambda p: p.stat().st_mtime)


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _secrets() -> List[Tuple[str, str]]:
    """설정에 저장된 판매자 정보·비밀값(패키지에 들어가면 안 되는 글자)."""
    try:
        import sp_api
        c = sp_api.load_config()
    except Exception:
        return []
    keys = ("seller_id", "report_email", "brand", "client_id", "client_secret", "refresh_token",
            "ads_client_id", "ads_client_secret", "ads_refresh_token", "ads_profile_id")
    return [(k, str(c.get(k) or "").strip()) for k in keys if len(str(c.get(k) or "").strip()) >= 4]


def check(fl: List[Path]) -> List[str]:
    bad = []
    sec = _secrets()
    for p in fl:
        try:
            t = p.read_bytes().decode("utf-8", errors="ignore").lower()
        except Exception:
            continue
        for k, v in sec:
            if v.lower() in t:
                bad.append(f"{p.relative_to(ROOT)}: 설정의 {k} 값이 들어 있음")
    return bad


def _manifest(entries: List[Tuple[str, bytes]], ver: str, released: str) -> str:
    lines = [f"{NAME} v{ver}", f"버전 날짜: {released}", f"패키지 만든 시각: {datetime.now():%Y-%m-%d %H:%M}",
             f"파일 {len(entries)}개", "", "크기\tSHA-256\t파일"]
    lines += [f"{len(d)}\t{_sha(d)}\t{n}" for n, d in entries]
    return "\n".join(lines) + "\n"


def _read_manifest(z: zipfile.ZipFile) -> Dict[str, str]:
    try:
        txt = z.read("RELEASE.txt").decode("utf-8")
    except KeyError:
        return {}
    out = {}
    for ln in txt.splitlines():
        parts = ln.split("\t")
        if len(parts) == 3 and parts[0].isdigit():
            out[parts[2]] = parts[1]
    return out


def build() -> Optional[Path]:
    sys.path.insert(0, str(BASE))
    import version
    ver = version.VERSION
    fl = files()
    bad = check(fl)
    if bad:
        print("❌ 패키지를 만들지 않았습니다. 판매자 정보·비밀값이 파일에 들어 있습니다:")
        for b in bad:
            print("   -", b)
        return None
    entries = [(p.relative_to(ROOT).as_posix(), p.read_bytes()) for p in fl]
    OUT.mkdir(parents=True, exist_ok=True)
    same = sorted(OUT.glob(f"{NAME}_v{ver}_*.zip"))
    if same:
        with zipfile.ZipFile(same[-1]) as z:
            old = _read_manifest(z)
        new = {n: _sha(d) for n, d in entries}
        if old == new:
            print(f"ℹ v{ver} 패키지가 이미 있고 내용도 같습니다: {same[-1].name}")
            return same[-1]
        changed = sorted(set(old) ^ set(new) | {n for n in new if n in old and old[n] != new[n]})
        print(f"❌ v{ver} 패키지가 이미 있는데 파일이 바뀌었습니다({len(changed)}개, 예: {', '.join(changed[:5])}).")
        print("   app/version.py에서 버전을 올린 뒤(고치기만 했으면 PATCH +1) 다시 만드세요.")
        return None
    path = OUT / f"{NAME}_v{ver}_{version.RELEASED}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in entries:
            z.writestr(n, d)
        z.writestr("RELEASE.txt", _manifest(entries, ver, version.RELEASED))
    print(f"✅ 패키지를 만들었습니다: releases/{path.name}  (파일 {len(entries)}개, {path.stat().st_size / 1024:,.0f} KB)")
    return path


def list_releases() -> List[Path]:
    rs = _glob("*.zip")
    if not rs:
        print("만든 패키지가 없습니다.")
    for p in rs:
        print(f"  {p.name}   {p.stat().st_size / 1024:,.0f} KB   {datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M}")
    return rs


def restore(ver: str) -> bool:
    ver = ver.lstrip("vV")
    cand = _glob(f"{ver}_*.zip")
    if not cand:
        print(f"❌ v{ver} 패키지가 없습니다.")
        list_releases()
        return False
    src = cand[-1]
    with zipfile.ZipFile(src) as z:
        names = [n for n in z.namelist() if n != "RELEASE.txt"]
        man = _read_manifest(z)
        for n in names:          # 패키지가 손상되지 않았는지 먼저 확인
            if man and _sha(z.read(n)) != man.get(n):
                print(f"❌ 패키지 파일이 손상되었습니다: {n}")
                return False
            if n.startswith("/") or ".." in Path(n).parts:
                print(f"❌ 올바르지 않은 파일 경로: {n}")
                return False
        # 지금 파일을 먼저 묶어 둠
        bdir = OUT / "backup"
        bdir.mkdir(parents=True, exist_ok=True)
        bak = bdir / f"before_restore_{datetime.now():%Y%m%d_%H%M%S}.zip"
        cur = files()
        with zipfile.ZipFile(bak, "w", zipfile.ZIP_DEFLATED) as b:
            for p in cur:
                b.write(p, p.relative_to(ROOT).as_posix())
        # 패키지에 없는 지금 프로그램 파일(app/*.py 등)은 지움 — 그 버전과 똑같게
        keep = set(names)
        for p in cur:
            rel = p.relative_to(ROOT).as_posix()
            if rel.startswith("app/") and rel not in keep:
                p.unlink()
        for n in names:
            dst = ROOT / n
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_bytes(z.read(n))
    print(f"✅ v{ver}로 되돌렸습니다({src.name}). 되돌리기 전 파일: releases/backup/{bak.name}")
    print("   reports(판매 데이터)·설정·로그인 정보는 그대로입니다. 화면 프로그램을 다시 여세요.")
    return True


def main(argv: List[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if len(argv) > 1 and argv[1] == "list":
        list_releases()
        return 0
    if len(argv) > 2 and argv[1] == "restore":
        return 0 if restore(argv[2]) else 1
    return 0 if build() else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
