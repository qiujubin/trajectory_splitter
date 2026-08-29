"""命令行入口 (调试 / 自动化用, GUI 走 gui.py)。"""

from __future__ import annotations

import argparse

from trajectory_splitter import (
    default_end_poi,
    default_start_poi,
    find_main_date,
    merge_two,
    split_at_poi,
)
from trajectory_splitter import logging_setup
from trajectory_splitter.modifier import ModifyOptions, modify_xls


def main() -> int:
    logging_setup.configure()
    parser = argparse.ArgumentParser(description="trajectory_splitter CLI")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_split = sub.add_parser(
        "split", help="单文件切片: 按 (start_poi, end_poi) 切成日前/后半段",
    )
    p_split.add_argument("src", help="输入 .xls 路径")
    p_split.add_argument(
        "--direction", choices=["日前", "后半"], required=True,
        help="日前 = 前半段, 后半 = 后半段",
    )
    p_split.add_argument("--out", required=True, help="输出目录")
    p_split.add_argument(
        "--start-poi", default=None,
        help=f'日前段起点 POI, 留空 = 从文件头切 (默认: {default_start_poi("日前")!r}); '
             f'后半段起点 POI (默认: {default_start_poi("后半")!r})',
    )
    p_split.add_argument(
        "--end-poi", default=None,
        help=f'前半段终点 POI (默认: {default_end_poi("日前")!r}); '
             f'后半段终点 POI (默认: {default_end_poi("后半")!r})',
    )
    p_split.add_argument(
        "--noon-cutoff", default="12:00:00",
        help='后半段"最后一次 start_poi"的时间上限, HH:MM:SS, 空字符串不限',
    )
    p_split.add_argument(
        "--no-clamp", dest="clamp_temp_to_0to4", action="store_false",
        help="不夹紧温度到 [0,4] (默认会夹紧)",
    )
    p_split.add_argument(
        "--filename-template", default="1{direction}.xls",
        help="输出文件名模板, 支持 {direction}",
    )
    p_split.add_argument(
        "--round-to-minute", action="store_true",
        help="时间列秒清零, 同分钟内重复行丢弃",
    )
    p_split.add_argument(
        "--dry-run", action="store_true",
        help="只看会输出到哪里 / 多少行, 不真写文件 (输出目录若不存在也不会建)",
    )

    p_merge = sub.add_parser(
        "merge", help="双文件拼接: 日前 + 后半 → 文件名按参考文件夹命名",
    )
    p_merge.add_argument("first", help="日前 .xls")
    p_merge.add_argument("second", help="后半 .xls")
    p_merge.add_argument("--out", required=True, help="输出目录")
    p_merge.add_argument(
        "--time-gap-minutes", type=int, default=1,
        help="日后第一行新时间相对日前最后一条的间隔 (分钟)",
    )
    p_merge.add_argument(
        "--no-clamp", dest="clamp_temp_to_0to4", action="store_false",
        help="不夹紧温度到 [0,4]",
    )
    p_merge.add_argument(
        "--filename-template", default="1日前.xls",
        help="输出文件名模板, 支持 {plate}",
    )
    p_merge.add_argument(
        "--output-plate", default=None,
        help="强制输出车牌 (覆盖前后两份所有行 + {plate}); 不传 = 用日前首行的车牌",
    )
    p_merge.add_argument(
        "--round-to-minute", action="store_true",
        help="时间列秒清零, 同分钟内重复行丢弃 (最后一步)",
    )
    p_merge.add_argument(
        "--dry-run", action="store_true",
        help="只看会输出到哪里 / 多少行, 不真写文件",
    )

    p_md = sub.add_parser("main-date", help="输出文件里记录数最多的那一天")
    p_md.add_argument("src", help="输入 .xls 路径")

    p_mod = sub.add_parser(
        "modify", help="数据微改: 车牌覆盖 + 时间偏移/微抖 + 温度微变",
    )
    p_mod.add_argument("src", help="输入 .xls 路径")
    p_mod.add_argument("--out", required=True, help="输出目录")
    p_mod.add_argument(
        "--output-plate", default="",
        help="输出车牌 (留空 = 不改; 填了就整列覆盖)",
    )
    p_mod.add_argument(
        "--round-to-minute", action="store_true",
        help="时间整点化: 秒清零 + 同分钟重复行丢弃",
    )
    p_mod.add_argument(
        "--time-offset-minutes", type=int, default=0,
        help="时间整体偏移 (分钟, 可正可负)",
    )
    p_mod.add_argument(
        "--time-jitter-seconds", type=float, default=8.0,
        help="每行时间正态微抖 σ (秒), μ=0",
    )
    p_mod.add_argument(
        "--temp-jitter", type=float, default=0.3,
        help="每行温度均匀随机 ±Y (度)",
    )
    p_mod.add_argument(
        "--no-clamp", dest="clamp_temp_to_0to4", action="store_false",
        help="不夹紧温度到 [0,4]",
    )
    p_mod.add_argument(
        "--filename-template", default="1modified.xls",
        help="输出文件名模板, 支持 {plate}",
    )
    p_mod.add_argument("--seed", type=int, default=None, help="随机种子 (复现用)")
    p_mod.add_argument(
        "--dry-run", dest="dry_run", action="store_true",
        help="只看会输出到哪里 / 多少行, 不真写文件",
    )

    args = parser.parse_args()

    if args.cmd == "split":
        noon = args.noon_cutoff or None
        out_path = split_at_poi(
            args.src, args.direction, args.out,
            start_poi=args.start_poi,
            end_poi=args.end_poi,
            noon_cutoff=noon,
            clamp_temp_to_0to4=args.clamp_temp_to_0to4,
            filename_template=args.filename_template,
            round_to_minute=args.round_to_minute,
            dry_run=args.dry_run,
        )
        prefix = "[DRY-RUN] 将输出: " if args.dry_run else "已输出: "
        print(f"{prefix}{out_path}")
        return 0

    if args.cmd == "merge":
        out_path = merge_two(
            args.first, args.second, args.out,
            time_gap_minutes=args.time_gap_minutes,
            clamp_temp_to_0to4=args.clamp_temp_to_0to4,
            filename_template=args.filename_template,
            output_plate=args.output_plate,
            round_to_minute=args.round_to_minute,
            dry_run=args.dry_run,
        )
        prefix = "[DRY-RUN] 将输出: " if args.dry_run else "已输出: "
        print(f"{prefix}{out_path}")
        return 0

    if args.cmd == "main-date":
        date, count = find_main_date(args.src)
        print(f"{date}\t{count}")
        return 0

    if args.cmd == "modify":
        opts = ModifyOptions(
            time_offset_minutes=args.time_offset_minutes,
            time_jitter_seconds=args.time_jitter_seconds,
            temp_jitter=args.temp_jitter,
            output_plate=args.output_plate,
            round_to_minute=args.round_to_minute,
            clamp_temp_to_0to4=args.clamp_temp_to_0to4,
            filename_template=args.filename_template,
            seed=args.seed,
            dry_run=args.dry_run,
        )
        out_path = modify_xls(args.src, args.out, opts)
        prefix = "[DRY-RUN] 将输出: " if args.dry_run else "已输出: "
        print(f"{prefix}{out_path}")
        return 0

    return 1


if __name__ == "__main__":
    raise SystemExit(main())