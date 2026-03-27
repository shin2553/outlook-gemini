"""
제품 키워드 → 매뉴얼 파일 매핑 인덱스
search_manuals(query) 함수로 관련 txt 파일 목록을 반환합니다.
"""

from pathlib import Path

EXPORT_DIR = Path(__file__).parent / "Manuals_txt"

# 키워드 → 검색할 txt 파일 목록
INDEX = {
    # ── 소프트웨어 ──────────────────────────────────────────────
    "laserdesk": ["laserDESK_manual.txt"],
    "laser desk": ["laserDESK_manual.txt"],
    "remote control": ["laserDESK_RemoteControl.en.txt"],
    "samlight": ["sc_SAMLight_en.txt"],
    "sam light": ["sc_SAMLight_en.txt"],
    "isCANcfg": ["Manual_iSCANcfg_v1-7.txt"],
    "iscancfg": ["Manual_iSCANcfg_v1-7.txt"],
    "rtc6conf": ["RTC6conf.txt"],
    "correction file": ["CorrectionFileConverter_Manual.txt"],

    # ── RTC 제어보드 ─────────────────────────────────────────────
    "rtc4": ["Manual_RTC4_Doc.1.9.1_en.txt"],
    "rtc5": ["Manual_RTC5_Doc.1.15.2_en.txt"],
    "rtc6": [
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",   # 최신
        "RTC6_Manual.en.txt",                           # v1.19.1 (구고객 호환)
    ],
    "rtc 6": [
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
        "RTC6_Manual.en.txt",
    ],
    "rtc board": [
        "Manual_RTC5_Doc.1.15.2_en.txt",
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
    ],
    "ethernet board": [
        "Manual_RTC4_Doc.1.9.1_en.txt",
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
    ],

    # ── excelliSCAN ──────────────────────────────────────────────
    "excelliscan 14": [
        "Manual_143094_excelliSCAN_14_Doc.3.2_en.txt",   # 1064nm
        "Manual_143416_excelliSCAN_14_Doc.3.0_en.txt",   # 343nm
    ],
    "excelliscan 20": ["Manual_144010_excelliSCAN_20_Doc.3.3_en.txt"],
    "excelliscan": [
        "Manual_143094_excelliSCAN_14_Doc.3.2_en.txt",
        "Manual_143416_excelliSCAN_14_Doc.3.0_en.txt",
        "Manual_144010_excelliSCAN_20_Doc.3.3_en.txt",
        "Manual_RTC6+excelliSCAN_Doc.1.0.11_en.txt",
        "excelliSCAN Serie_with_SCANahead-Technology.txt",
    ],
    "scanahead": [
        "Manual_RTC6+excelliSCAN_Doc.1.0.11_en.txt",
        "excelliSCAN Serie_with_SCANahead-Technology.txt",
    ],

    # ── intelliSCAN ──────────────────────────────────────────────
    "intelliscan se 10": [
        "Manual_124445_intelliSCAN_se_10_Doc.2.6_en.txt",  # 355nm
        "Manual_148914_intelliSCAN_se_10_Doc.2.6_en.txt",  # 1030nm
    ],
    "intelliscan se 14": [
        "Manual_127242_intelliSCAN_se_14_Doc.2.6_en.txt",  # 1030nm
        "Manual_131606_intelliSCAN_se_14_Doc.2.6_en.txt",  # 355nm
        "Manual_146907_intelliSCAN_se_14_Doc.2.6_en.txt",  # 266nm
    ],
    "intelliscan se": [
        "Manual_124445_intelliSCAN_se_10_Doc.2.6_en.txt",
        "Manual_148914_intelliSCAN_se_10_Doc.2.6_en.txt",
        "Manual_127242_intelliSCAN_se_14_Doc.2.6_en.txt",
        "Manual_131606_intelliSCAN_se_14_Doc.2.6_en.txt",
        "Manual_146907_intelliSCAN_se_14_Doc.2.6_en.txt",
    ],
    "intelliscan iii 10": ["Manual_151262_intelliSCAN_III_10_Module_2022-09-11_en.txt"],
    "intelliscan iii 30": ["Manual_133952_intelliSCAN_III_30_Doc.2.6_en.txt"],
    "intelliscan iii": [
        "Manual_151262_intelliSCAN_III_10_Module_2022-09-11_en.txt",
        "Manual_133952_intelliSCAN_III_30_Doc.2.6_en.txt",
    ],
    "intelliscan 10": ["Manual_138459_intelliSCAN_10_Doc.2.6_en.txt"],
    "intelliscan 14": ["Manual_140473_intelliSCAN_14_Doc.2.6_en.txt"],
    "intelliscan 20": [
        "Manual_142911_intelliSCAN_20_Doc.2.6_en.txt",   # 1055~1085nm
        "Manual_155060_intelliSCAN_20_Doc.2.6_en.txt",   # 515+532nm
    ],
    "intelliscan 30": ["Manual_151197_intelliSCAN_30_Doc.2.6_en.txt"],
    "intelliscan": [
        "Manual_138459_intelliSCAN_10_Doc.2.6_en.txt",
        "Manual_140473_intelliSCAN_14_Doc.2.6_en.txt",
        "Manual_142911_intelliSCAN_20_Doc.2.6_en.txt",
        "Manual_155060_intelliSCAN_20_Doc.2.6_en.txt",
        "Manual_151197_intelliSCAN_30_Doc.2.6_en.txt",
        "Manual_133952_intelliSCAN_III_30_Doc.2.6_en.txt",
        "Manual_151262_intelliSCAN_III_10_Module_2022-09-11_en.txt",
        "Manual_124445_intelliSCAN_se_10_Doc.2.6_en.txt",
        "Manual_148914_intelliSCAN_se_10_Doc.2.6_en.txt",
        "Manual_127242_intelliSCAN_se_14_Doc.2.6_en.txt",
        "Manual_131606_intelliSCAN_se_14_Doc.2.6_en.txt",
        "Manual_146907_intelliSCAN_se_14_Doc.2.6_en.txt",
    ],
    "idrive": [   # iDRIVE 관련 질문 → intelliSCAN/excelliSCAN 전체
        "Manual_138459_intelliSCAN_10_Doc.2.6_en.txt",
        "Manual_140473_intelliSCAN_14_Doc.2.6_en.txt",
        "Manual_143094_excelliSCAN_14_Doc.3.2_en.txt",
        "Manual_143416_excelliSCAN_14_Doc.3.0_en.txt",
        "Manual_iSCANcfg_v1-7.txt",
    ],

    # ── SCANcube ─────────────────────────────────────────────────
    "scancube iv 14": [
        "Manual_SCANcube_IV_14_XY2-100_Part1_Doc.1.0_en.txt",
        "Manual_148094_SCANcube_IV_14_Part2_Doc.2.3_en.txt",
        "Pin-out_SCANcubeIV_en.txt",
    ],
    "scancube iii": ["Manual_136789_SCANcube_III_10_Doc.2.6_en.txt"],
    "scancube 10": [
        "Manual_112612_SCANcube_10_Doc.2.6_en.txt",
        "Manual_122369_SCANcube_10_Doc.2.6_en.txt",
    ],
    "scancube": [
        "Manual_112612_SCANcube_10_Doc.2.6_en.txt",
        "Manual_122369_SCANcube_10_Doc.2.6_en.txt",
        "Manual_136789_SCANcube_III_10_Doc.2.6_en.txt",
        "Manual_SCANcube_IV_14_XY2-100_Part1_Doc.1.0_en.txt",
        "Manual_148094_SCANcube_IV_14_Part2_Doc.2.3_en.txt",
    ],
    "basicube": ["Manual_135990_basiCube_10_Doc.1.13_en.txt"],
    "hurriscan": ["Manual_114179_hurrySCAN_II_7_Doc.2.6_en.txt"],
    "hurryscan": ["Manual_114179_hurrySCAN_II_7_Doc.2.6_en.txt"],
    "varioscan": [
        "Manual_varioSCAN_de_II_20i_Part1_Doc.1.1.8_en.txt",
        "Manual_151300_varioSCAN_de_II_20i_FT_Part2_Doc.1.1.4_en.txt",
    ],

    # ── 인터페이스 & 주변기기 ────────────────────────────────────
    "dsib": [
        "Manual_115648_DSIB_Interface_Board_2019-03-20_en.txt",
        "Manual_117462_DSIB_SL2-100_Interface_Board_2022-03-01_en.txt",
    ],
    "sl2-100": [
        "Manual_117462_DSIB_SL2-100_Interface_Board_2022-03-01_en.txt",
    ],
    "xy2-100": [
        "Datasheet_0125377_XY2-100_Converter_II_2025-04-29_en.txt",
        "Manual_SCANcube_IV_14_XY2-100_Part1_Doc.1.0_en.txt",
    ],
    "da40": ["Datasheet_104553_DA40_Adapter_Board_2017-09-13_en.txt"],
    "camera": ["Manual_127502_Cameraadapter_Doc.1.6_en.txt"],
    "single axis": ["Manual_153577 153578_Single-Axis-Module_Doc.2.2_en.txt"],
    "dynaxis": ["Manual_153577 153578_Single-Axis-Module_Doc.2.2_en.txt"],

    # ── 기술문서 ─────────────────────────────────────────────────
    "3axis": ["AppNote_Calibrating_3-Axis_Laser_Scan_System_Doc.1.5.0_en.txt"],
    "3d": ["AppNote_Calibrating_3-Axis_Laser_Scan_System_Doc.1.5.0_en.txt",
           "laserDESK_manual.txt"],
    "calibrat": ["AppNote_Calibrating_3-Axis_Laser_Scan_System_Doc.1.5.0_en.txt"],
    "mof": ["Introduction_to_MoF_with_excelliSCAN_FOR_KOREA_(by_MSL_2022_Rev._01).txt"],
    "marking on the fly": ["Introduction_to_MoF_with_excelliSCAN_FOR_KOREA_(by_MSL_2022_Rev._01).txt"],

    # ── 공통 기능 키워드 ─────────────────────────────────────────
    "logging": ["laserDESK_manual.txt"],
    "protocol": ["laserDESK_manual.txt"],
    "actual position": [
        "laserDESK_manual.txt",
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
    ],
    "galvo": [
        "laserDESK_manual.txt",
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
    ],
    "scan head": [
        "laserDESK_manual.txt",
        "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt",
    ],
    "laser parameter": ["laserDESK_manual.txt"],
    "marking": ["laserDESK_manual.txt"],
    "job": ["laserDESK_manual.txt"],
    "variant": ["laserDESK_manual.txt"],
    "barcode": ["laserDESK_manual.txt"],
    "serial number": ["laserDESK_manual.txt"],
    "wobble": ["laserDESK_manual.txt"],
    "skywriting": ["laserDESK_manual.txt"],
    "fill": ["laserDESK_manual.txt"],
    "hatch": ["laserDESK_manual.txt"],
    "speed": ["laserDESK_manual.txt", "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt"],
    "delay": ["laserDESK_manual.txt", "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt"],
    "power": ["laserDESK_manual.txt"],
    "frequency": ["laserDESK_manual.txt"],
    "trigger": ["laserDESK_manual.txt", "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt"],
    "encoder": ["laserDESK_manual.txt", "Manual_RTC6_SW-V1.22.0_Doc.1.1.3_en.txt"],
    "remote": ["laserDESK_RemoteControl.en.txt"],
}


def search_manuals(query: str) -> list[Path]:
    """
    문의 텍스트에서 관련 매뉴얼 txt 파일 경로 목록을 반환합니다.
    1) INDEX 키워드 매칭
    2) INDEX 미등록 신규 파일 자동 탐지 (파일명 기반)
    중복 제거 후 존재하는 파일만 반환합니다.
    """
    q = query.lower()
    matched = []
    seen = set()

    # 1. INDEX 기반 검색
    for keyword, files in INDEX.items():
        if keyword in q:
            for f in files:
                if f not in seen:
                    path = EXPORT_DIR / f
                    if path.exists():
                        matched.append(path)
                        seen.add(f)

    # 2. INDEX 미등록 신규 파일 자동 탐지
    all_indexed = {f for files in INDEX.values() for f in files}
    q_words = [w for w in q.split() if len(w) > 1]
    for txt_file in sorted(EXPORT_DIR.glob("*.txt")):
        if txt_file.name not in all_indexed and txt_file.name not in seen:
            file_keywords = txt_file.name.lower().replace("_", " ").replace("-", " ").replace(".", " ")
            if any(w in file_keywords for w in q_words):
                matched.append(txt_file)
                seen.add(txt_file.name)

    # 3. 아무것도 매칭 안 되면 laserDESK 기본 매뉴얼 반환
    if not matched:
        default = EXPORT_DIR / "laserDESK_manual.txt"
        if default.exists():
            matched.append(default)

    return matched


if __name__ == "__main__":
    # 테스트
    tests = [
        "excelliscan14 + rtc6 로깅 문제",
        "intelliscan se 14 actual position",
        "samlight에서 rtc5 연결하는 방법",
        "scancube iv 14 xy2-100 설정",
    ]
    for t in tests:
        results = search_manuals(t)
        print(f"\n[{t}]")
        for r in results:
            print(f"  → {r.name}")
