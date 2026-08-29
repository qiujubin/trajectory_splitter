"""Tkinter GUI: 两个 Tab — 单文件切片 (按 start_poi/end_poi) + 双文件拼接。

启动:  python -m trajectory_splitter.gui
"""

from __future__ import annotations

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from trajectory_splitter import (
    default_end_poi,
    default_start_poi,
    find_main_date,
    merge_two,
    modify_xls,
    split_at_poi,
)
from trajectory_splitter import logging_setup
from trajectory_splitter.gui_state import load_state, save_state
from trajectory_splitter.modifier import ModifyOptions


# 把 GUI 变量集中登记, 方便 load/save 一把梭
SPLIT_KEYS: list[str] = [
    "split_src", "split_dir", "split_start_poi", "split_end_poi",
    "split_noon", "split_clamp", "split_round_to_minute",
    "split_tpl", "split_out",
]
MERGE_KEYS: list[str] = [
    "merge_first", "merge_second", "merge_gap",
    "merge_clamp", "merge_round_to_minute",
    "merge_tpl", "merge_out",
]
MODIFY_KEYS: list[str] = [
    "mod_src", "mod_out_plate",
    "mod_time_offset", "mod_time_jitter", "mod_temp_jitter",
    "mod_clamp", "mod_round_to_minute",
    "mod_tpl", "mod_out", "mod_seed",
]


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("trajectory_splitter")
        root.geometry("880x760")
        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_split = ttk.Frame(nb)
        self.tab_merge = ttk.Frame(nb)
        self.tab_modify = ttk.Frame(nb)
        nb.add(self.tab_split, text="单文件切片 (按 POI)")
        nb.add(self.tab_merge, text="双文件拼接")
        nb.add(self.tab_modify, text="数据微改")
        self._build_split_tab(self.tab_split)
        self._build_merge_tab(self.tab_merge)
        self._build_modify_tab(self.tab_modify)
        # load 持久化状态 (在所有 StringVar/BooleanVar 创建完之后)
        self._restore_state()
        # 关窗时 save
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- 持久化 ----------------
    def _restore_state(self) -> None:
        state = load_state()
        for sect, keys in (("split", SPLIT_KEYS), ("merge", MERGE_KEYS), ("modify", MODIFY_KEYS)):
            sect_state = state.get(sect) or {}
            for key in keys:
                if key in sect_state:
                    getattr(self, key).set(sect_state[key])

    def _snapshot_state(self) -> dict[str, dict[str, object]]:
        return {
            "split": {k: getattr(self, k).get() for k in SPLIT_KEYS},
            "merge": {k: getattr(self, k).get() for k in MERGE_KEYS},
            "modify": {k: getattr(self, k).get() for k in MODIFY_KEYS},
        }

    def _on_close(self) -> None:
        save_state(self._snapshot_state())
        self.root.destroy()

    # ---------------- 单文件切片 ----------------
    def _build_split_tab(self, parent: ttk.Frame) -> None:
        f = ttk.Frame(parent, padding=12)
        f.pack(fill="x")

        self.split_src = tk.StringVar()
        self.split_dir = tk.StringVar(value="日前")
        self.split_start_poi = tk.StringVar()
        self.split_end_poi = tk.StringVar()
        self.split_noon = tk.StringVar(value="12:00:00")
        self.split_clamp = tk.BooleanVar(value=True)
        self.split_round_to_minute = tk.BooleanVar(value=False)
        self.split_tpl = tk.StringVar(value="1{direction}.xls")
        self.split_out = tk.StringVar()

        # 初始化 POI 默认值 (按当前方向)
        self._refresh_split_poi_defaults()

        ttk.Label(f, text="输入 .xls:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.split_src, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(f, text="选文件", command=self._pick_split_src).grid(row=0, column=2)

        ttk.Label(f, text="方向 (日前=前半段, 后半=后半段):").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        self.split_dir_combo = ttk.Combobox(
            f, textvariable=self.split_dir, values=["日前", "后半"],
            state="readonly", width=8,
        )
        self.split_dir_combo.grid(row=1, column=1, sticky="w", padx=4)
        self.split_dir_combo.bind(
            "<<ComboboxSelected>>", lambda _e: self._refresh_split_poi_defaults(),
        )

        ttk.Label(f, text="start_poi (日前留空=从文件头; 后半=麻涌卸货点):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.split_start_poi, width=25).grid(
            row=2, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="end_poi (日前=麻涌卸货点; 后半=广州朴朴):").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.split_end_poi, width=25).grid(
            row=3, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="noon_cutoff (后半段『最后一次 start_poi』时间上限, 空=不限):").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.split_noon, width=12).grid(
            row=4, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="夹紧温度到 [0,4] (温和修正, 不用随机数):").grid(
            row=5, column=0, sticky="w", pady=4,
        )
        ttk.Checkbutton(f, variable=self.split_clamp).grid(
            row=5, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="整点化 (秒清零, 同分钟重复行丢弃):").grid(
            row=6, column=0, sticky="w", pady=4,
        )
        ttk.Checkbutton(f, variable=self.split_round_to_minute).grid(
            row=6, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="输出文件名模板 (用 {direction} / {plate} / {date}):").grid(
            row=7, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.split_tpl, width=30).grid(
            row=7, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="输出目录:").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.split_out, width=55).grid(row=8, column=1, padx=4)
        ttk.Button(f, text="选目录", command=self._pick_split_out).grid(row=8, column=2)

        ttk.Button(f, text="预览主要日期", command=self._preview_split).grid(
            row=9, column=1, sticky="w", pady=8,
        )
        ttk.Button(f, text="运行", command=self._run_split).grid(
            row=9, column=1, sticky="e", pady=8,
        )

        self.split_info = ttk.Label(f, text="(未选文件)", foreground="gray")
        self.split_info.grid(row=10, column=1, sticky="w")

        log_frame = ttk.LabelFrame(parent, text="日志")
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.split_log = tk.Text(log_frame, height=8)
        self.split_log.pack(fill="both", expand=True)

    def _refresh_split_poi_defaults(self) -> None:
        """切方向时强制刷新 POI 默认值。

        之前的逻辑是"为空才填", 但首次构建时已经填过了默认值,
        切方向时空值检查会失败 -> 用户看到的还是上一次的默认值。
        切方向是显式动作, 应该无脑刷成新方向的默认值。
        """
        direction = self.split_dir.get() or "日前"
        self.split_start_poi.set(default_start_poi(direction))
        self.split_end_poi.set(default_end_poi(direction))

    def _pick_split_src(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.split_src.set(path)

    def _pick_split_out(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.split_out.set(path)

    def _preview_split(self) -> None:
        src = self.split_src.get().strip()
        if not src:
            messagebox.showwarning("提示", "先选输入文件")
            return
        try:
            main_date, count = find_main_date(src)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("读文件失败", str(e))
            return
        msg = f"主要日期: {main_date}\n该日记录数: {count}"
        self.split_info.config(
            text=f"主要日期 = {main_date} ({count} 条)", foreground="blue",
        )
        messagebox.showinfo("主要日期", msg)

    def _run_split(self) -> None:
        src = self.split_src.get().strip()
        direction = self.split_dir.get()
        start_poi = self.split_start_poi.get().strip()
        end_poi = self.split_end_poi.get().strip()
        noon = self.split_noon.get().strip() or None
        clamp = self.split_clamp.get()
        round_to_minute = self.split_round_to_minute.get()
        tpl = self.split_tpl.get().strip() or "1{direction}.xls"
        out_dir = self.split_out.get().strip()
        if not all([src, out_dir]):
            messagebox.showwarning("提示", "源文件 / 输出目录都得填")
            return
        if not end_poi:
            messagebox.showwarning("提示", "end_poi 不能为空")
            return
        if direction == "后半" and not start_poi:
            messagebox.showwarning("提示", "direction='后半' 必须指定 start_poi (后半段起点)")
            return
        try:
            out_path = split_at_poi(
                src, direction, out_dir,
                start_poi=start_poi,
                end_poi=end_poi,
                noon_cutoff=noon,
                clamp_temp_to_0to4=clamp,
                filename_template=tpl,
                round_to_minute=round_to_minute,
            )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("切片失败", str(e))
            self._log_split(f"[ERR] {e}")
            return
        msg = f"已输出: {out_path}"
        self._log_split(msg)
        messagebox.showinfo("完成", msg)

    def _log_split(self, msg: str) -> None:
        self.split_log.insert("end", msg + "\n")
        self.split_log.see("end")

    # ---------------- 双文件拼接 ----------------
    def _build_merge_tab(self, parent: ttk.Frame) -> None:
        f = ttk.Frame(parent, padding=12)
        f.pack(fill="x")

        self.merge_first = tk.StringVar()
        self.merge_second = tk.StringVar()
        self.merge_gap = tk.StringVar(value="1")
        self.merge_out_plate = tk.StringVar()
        self.merge_clamp = tk.BooleanVar(value=True)
        self.merge_round_to_minute = tk.BooleanVar(value=False)
        self.merge_tpl = tk.StringVar(value="1日前.xls")
        self.merge_out = tk.StringVar()

        ttk.Label(f, text="前半 (日前) .xls:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.merge_first, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(f, text="选文件", command=self._pick_merge_first).grid(row=0, column=2)

        ttk.Label(f, text="后半 .xls:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.merge_second, width=55).grid(row=1, column=1, padx=4)
        ttk.Button(f, text="选文件", command=self._pick_merge_second).grid(row=1, column=2)

        ttk.Label(f, text="时间间隔 (分钟, 后半首条相对前半末条):").grid(
            row=2, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.merge_gap, width=10).grid(
            row=2, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="输出车牌 (留空 = 用前半首行的车牌):").grid(
            row=3, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.merge_out_plate, width=20).grid(
            row=3, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="夹紧温度到 [0,4] (温和修正, 不用随机数):").grid(
            row=4, column=0, sticky="w", pady=4,
        )
        ttk.Checkbutton(f, variable=self.merge_clamp).grid(
            row=4, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="整点化 (秒清零, 同分钟重复行丢弃):").grid(
            row=5, column=0, sticky="w", pady=4,
        )
        ttk.Checkbutton(f, variable=self.merge_round_to_minute).grid(
            row=5, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="输出文件名模板 (用 {plate} / {date}):").grid(
            row=6, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.merge_tpl, width=30).grid(
            row=6, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="输出目录:").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.merge_out, width=55).grid(row=7, column=1, padx=4)
        ttk.Button(f, text="选目录", command=self._pick_merge_out).grid(row=7, column=2)

        ttk.Button(f, text="运行", command=self._run_merge).grid(
            row=8, column=1, sticky="e", pady=8,
        )

        log_frame = ttk.LabelFrame(parent, text="日志")
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.merge_log = tk.Text(log_frame, height=8)
        self.merge_log.pack(fill="both", expand=True)

    def _pick_merge_first(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.merge_first.set(path)

    def _pick_merge_second(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.merge_second.set(path)

    def _pick_merge_out(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.merge_out.set(path)

    def _run_merge(self) -> None:
        first = self.merge_first.get().strip()
        second = self.merge_second.get().strip()
        gap_str = self.merge_gap.get().strip()
        out_plate = self.merge_out_plate.get().strip() or None
        clamp = self.merge_clamp.get()
        round_to_minute = self.merge_round_to_minute.get()
        tpl = self.merge_tpl.get().strip() or "1日前.xls"
        out_dir = self.merge_out.get().strip()
        if not all([first, second, gap_str, out_dir]):
            messagebox.showwarning("提示", "前半 / 后半 / 时间间隔 / 输出目录都得填")
            return
        try:
            gap = int(gap_str)
        except ValueError:
            messagebox.showerror("错误", "时间间隔必须是整数分钟")
            return
        try:
            out_path = merge_two(
                first, second, out_dir,
                time_gap_minutes=gap,
                clamp_temp_to_0to4=clamp,
                filename_template=tpl,
                output_plate=out_plate,
                round_to_minute=round_to_minute,
            )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("拼接失败", str(e))
            self._log_merge(f"[ERR] {e}")
            return
        msg = f"已输出: {out_path}"
        self._log_merge(msg)
        messagebox.showinfo("完成", msg)

    def _log_merge(self, msg: str) -> None:
        self.merge_log.insert("end", msg + "\n")
        self.merge_log.see("end")

    # ---------------- 数据微改 ----------------
    def _build_modify_tab(self, parent: ttk.Frame) -> None:
        f = ttk.Frame(parent, padding=12)
        f.pack(fill="x")

        self.mod_src = tk.StringVar()
        self.mod_out_plate = tk.StringVar()
        self.mod_time_offset = tk.StringVar(value="0")
        self.mod_time_jitter = tk.StringVar(value="8")
        self.mod_temp_jitter = tk.StringVar(value="0.3")
        self.mod_clamp = tk.BooleanVar(value=True)
        self.mod_round_to_minute = tk.BooleanVar(value=False)
        self.mod_tpl = tk.StringVar(value="1modified.xls")
        self.mod_out = tk.StringVar()
        self.mod_seed = tk.StringVar(value="")

        ttk.Label(f, text="输入 .xls:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.mod_src, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(f, text="选文件", command=self._pick_mod_src).grid(row=0, column=2)

        ttk.Label(f, text="输出车牌 (留空 = 不改; 填了就覆盖所有行):").grid(
            row=1, column=0, sticky="w", pady=4,
        )
        ttk.Entry(f, textvariable=self.mod_out_plate, width=20).grid(
            row=1, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="时间整体偏移 (分钟, 可正可负):").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_time_offset, width=10).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Label(f, text="时间微抖 σ (秒/行, 默认 8, 推荐 5-30):").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_time_jitter, width=10).grid(row=3, column=1, sticky="w", padx=4)

        ttk.Label(f, text="温度微抖 ±Y 度 (推荐 0.2-0.5, 默认 0.3):").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_temp_jitter, width=10).grid(row=4, column=1, sticky="w", padx=4)

        ttk.Label(f, text="夹紧温度到 [0,4] (走分段映射规则):").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Checkbutton(f, variable=self.mod_clamp).grid(row=5, column=1, sticky="w", padx=4)

        ttk.Label(f, text="整点化 (秒清零, 同分钟重复行丢弃):").grid(
            row=6, column=0, sticky="w", pady=4,
        )
        ttk.Checkbutton(f, variable=self.mod_round_to_minute).grid(
            row=6, column=1, sticky="w", padx=4,
        )

        ttk.Label(f, text="随机种子 (空=真随机):").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_seed, width=10).grid(row=7, column=1, sticky="w", padx=4)

        ttk.Label(f, text="输出文件名模板 (用 {plate} / {date}):").grid(row=8, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_tpl, width=30).grid(row=8, column=1, sticky="w", padx=4)

        ttk.Label(f, text="输出目录:").grid(row=9, column=0, sticky="w", pady=4)
        ttk.Entry(f, textvariable=self.mod_out, width=55).grid(row=9, column=1, padx=4)
        ttk.Button(f, text="选目录", command=self._pick_mod_out).grid(row=9, column=2)

        ttk.Button(f, text="运行", command=self._run_modify).grid(row=10, column=1, sticky="e", pady=8)

        log_frame = ttk.LabelFrame(parent, text="日志")
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.mod_log = tk.Text(log_frame, height=8)
        self.mod_log.pack(fill="both", expand=True)

    def _pick_mod_src(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel 97-2003", "*.xls")])
        if path:
            self.mod_src.set(path)

    def _pick_mod_out(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.mod_out.set(path)

    def _run_modify(self) -> None:
        src = self.mod_src.get().strip()
        out_dir = self.mod_out.get().strip()
        if not all([src, out_dir]):
            messagebox.showwarning("提示", "源文件 / 输出目录都得填")
            return
        try:
            offset = int(self.mod_time_offset.get().strip() or "0")
            jitter_s = float(self.mod_time_jitter.get().strip() or "8")
            jitter_t = float(self.mod_temp_jitter.get().strip() or "0.3")
        except ValueError:
            messagebox.showerror("错误", "时间偏移/微抖/温度微抖都必须是数字")
            return
        seed_raw = self.mod_seed.get().strip()
        seed = int(seed_raw) if seed_raw else None

        out_plate = self.mod_out_plate.get().strip()

        tpl = self.mod_tpl.get().strip() or "1modified.xls"

        opts = ModifyOptions(
            time_offset_minutes=offset,
            time_jitter_seconds=jitter_s,
            temp_jitter=jitter_t,
            output_plate=out_plate,
            round_to_minute=self.mod_round_to_minute.get(),
            clamp_temp_to_0to4=self.mod_clamp.get(),
            filename_template=tpl,
            seed=seed,
        )
        try:
            out_path = modify_xls(src, out_dir, opts)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("微改失败", str(e))
            self._log_mod(f"[ERR] {e}")
            return
        msg = f"已输出: {out_path}"
        self._log_mod(msg)
        messagebox.showinfo("完成", msg)

    def _log_mod(self, msg: str) -> None:
        self.mod_log.insert("end", msg + "\n")
        self.mod_log.see("end")


def main() -> int:
    logging_setup.configure()
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())