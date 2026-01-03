import ctypes
import json
import math
import os
import socket
import subprocess
import sys
import threading
import time
import tkinter as tk
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from tkinter import colorchooser, messagebox, ttk
#©️ 2026 LUCA.NEX


def _load_port_value():
    path = os.path.join(os.getcwd(), "port.txt")
    try:
        with open(path, "r", encoding="utf-8") as handle:
            raw = handle.read().strip()
        port = int(raw)
        if 1 <= port <= 65535:
            return port
    except Exception:
        pass
    return 8989


def _is_windows():
    return os.name == "nt"


def _is_admin():
    if not _is_windows():
        return False
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def _relaunch_as_admin():
    params = " ".join([f'"{arg}"' if " " in arg else arg for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)


def _ensure_firewall_rule(port):
    if not _is_windows():
        return
    if not _is_admin():
        _relaunch_as_admin()
        sys.exit(0)
    rule_name = f"MSFS 2020/2024 Flight Data Export {port}"
    subprocess.run(
        ["netsh", "advfirewall", "firewall", "delete", "rule", f"name={rule_name}"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    subprocess.run(
        [
            "netsh",
            "advfirewall",
            "firewall",
            "add",
            "rule",
            f"name={rule_name}",
            "dir=in",
            "action=allow",
            "protocol=TCP",
            f"localport={port}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


class DataStore:
    def __init__(self):
        self._lock = threading.Lock()
        self._data = {
            "altitude_ft": 0.0,
            "heading_deg": 0.0,
            "airspeed_kt": 0.0,
            "vertical_speed_fpm": 0.0,
            "latitude": 0.0,
            "longitude": 0.0,
            "pitch_deg": 0.0,
            "bank_deg": 0.0,
            "fuel_total_gal": 0.0,
            "last_update": time.time(),
            "source": "idle",
        }

    def update(self, new_data):
        with self._lock:
            self._data.update(new_data)
            self._data["last_update"] = time.time()

    def snapshot(self):
        with self._lock:
            return dict(self._data)


class DataServer:
    def __init__(self, host, port, data_provider, static_dir):
        self._host = host
        self._port = port
        self._data_provider = data_provider
        self._static_dir = static_dir
        self._server = None

    def start(self):
        if self._server:
            return
        handler = self._make_handler()
        self._server = ThreadingHTTPServer((self._host, self._port), handler)
        self._server.serve_forever()

    def stop(self):
        if not self._server:
            return
        self._server.shutdown()
        self._server.server_close()
        self._server = None

    def _make_handler(self):
        data_provider = self._data_provider
        static_dir = self._static_dir

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                if self.path == "/" or self.path.startswith("/index.html"):
                    self._send_file("index.html", "text/html; charset=utf-8")
                    return
                if self.path.startswith("/style.css"):
                    self._send_file("style.css", "text/css; charset=utf-8")
                    return
                if self.path.startswith("/html-lang.json"):
                    self._send_file("html-lang.json", "application/json; charset=utf-8")
                    return
                if self.path.startswith("/data"):
                    payload = json.dumps(data_provider(), ensure_ascii=False)
                    self._send_bytes(
                        payload.encode("utf-8"), "application/json; charset=utf-8"
                    )
                    return
                self.send_response(404)
                self.end_headers()

            def log_message(self, format, *args):
                return

            def _send_file(self, filename, content_type):
                path = os.path.join(static_dir, filename)
                try:
                    with open(path, "rb") as f:
                        body = f.read()
                except OSError:
                    self.send_response(404)
                    self.end_headers()
                    return
                self._send_bytes(body, content_type)

            def _send_bytes(self, body, content_type):
                self.send_response(200)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        return Handler


class SimConnectReader:
    def __init__(self):
        self._available = False
        self._connected = False
        self._simconnect = None
        self._requests = None

        try:
            from SimConnect import SimConnect, AircraftRequests

            self._SimConnect = SimConnect
            self._AircraftRequests = AircraftRequests
            self._available = True
        except Exception:
            self._available = False

    @property
    def available(self):
        return self._available

    def connect(self):
        if not self._available or self._connected:
            return self._connected
        try:
            self._simconnect = self._SimConnect()
            self._requests = self._AircraftRequests(self._simconnect, _time=2000)
            self._connected = True
        except Exception:
            self._connected = False
        return self._connected

    def read(self):
        if not self._connected:
            return None
        try:
            return {
                "altitude_ft": float(self._requests.get("PLANE_ALTITUDE")),
                "heading_deg": float(self._requests.get("PLANE_HEADING_DEGREES_TRUE")),
                "airspeed_kt": float(self._requests.get("AIRSPEED_INDICATED")),
                "vertical_speed_fpm": float(self._requests.get("VERTICAL_SPEED")),
                "latitude": float(self._requests.get("PLANE_LATITUDE")),
                "longitude": float(self._requests.get("PLANE_LONGITUDE")),
                "pitch_deg": float(self._requests.get("PLANE_PITCH_DEGREES")),
                "bank_deg": float(self._requests.get("PLANE_BANK_DEGREES")),
                "fuel_total_gal": float(self._requests.get("FUEL_TOTAL_QUANTITY")),
                "source": "simconnect",
            }
        except Exception:
            return None


class FsuipcReader:
    def __init__(self):
        self._available = False
        self._connected = False
        self._pyuipc = None
        self._prepared = None
        self._offset_specs = []

        try:
            import pyuipc

            self._pyuipc = pyuipc
            self._available = True
        except Exception:
            self._available = False

        self._offset_specs = self._load_offsets()

    @property
    def available(self):
        return self._available

    def connect(self):
        if not self._available or self._connected:
            return self._connected
        try:
            self._pyuipc.open(self._pyuipc.SIM_ANY)
            self._prepared = self._pyuipc.prepare_data(
                [(spec["offset"], spec["type"]) for _key, spec in self._offset_specs]
            )
            self._connected = True
        except Exception:
            self._connected = False
        return self._connected

    def read(self):
        if not self._connected:
            return None
        try:
            values = self._pyuipc.read(self._prepared)
        except Exception:
            self._connected = False
            return None
        data = {}
        for (key, spec), raw in zip(self._offset_specs, values):
            data[key] = self._convert_value(raw, spec)
        data["source"] = "fsuipc"
        return data

    def _convert_value(self, raw, spec):
        if isinstance(raw, (bytes, bytearray)):
            try:
                return raw.decode("utf-8", errors="ignore").strip("\x00")
            except Exception:
                return ""
        scale = float(spec.get("scale", 1.0))
        offset_add = float(spec.get("offset_add", 0.0))
        divisor = spec.get("divisor")
        value = float(raw)
        if divisor:
            value = value / float(divisor)
        return value * scale + offset_add

    def _load_offsets(self):
        path = os.path.join(os.getcwd(), "fsuipc_offsets.json")
        if not os.path.exists(path):
            self._write_default_offsets(path)
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            data = self._default_offsets()
        specs = []
        for key, spec in data.items():
            try:
                offset_value = int(str(spec["offset"]), 16)
            except Exception:
                continue
            specs.append(
                (
                    key,
                    {
                        "offset": offset_value,
                        "type": spec.get("type", "d"),
                        "scale": spec.get("scale", 1.0),
                        "offset_add": spec.get("offset_add", 0.0),
                        "divisor": spec.get("divisor"),
                    },
                )
            )
        return specs

    def _write_default_offsets(self, path):
        try:
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(self._default_offsets(), handle, indent=2)
        except Exception:
            pass

    def _default_offsets(self):
        return {
            "altitude_ft": {
                "offset": "0x0570",
                "type": "d",
                "scale": 1 / 65536,
            },
            "heading_deg": {
                "offset": "0x0580",
                "type": "u",
                "scale": 360 / 65536,
            },
            "airspeed_kt": {
                "offset": "0x02BC",
                "type": "u",
                "scale": 1 / 128,
            },
            "vertical_speed_fpm": {
                "offset": "0x02C8",
                "type": "d",
                "scale": 1 / 256,
            },
            "latitude": {
                "offset": "0x0560",
                "type": "l",
                "scale": 90 / 2147483648,
            },
            "longitude": {
                "offset": "0x0568",
                "type": "l",
                "scale": 360 / 4294967296,
            },
            "pitch_deg": {
                "offset": "0x0578",
                "type": "d",
                "scale": 360 / 65536,
            },
            "bank_deg": {
                "offset": "0x057C",
                "type": "d",
                "scale": 360 / 65536,
            },
            "fuel_total_gal": {
                "offset": "0x0B7C",
                "type": "u",
                "scale": 1.0,
            },
        }


class DataCollector(threading.Thread):
    def __init__(self, store):
        super().__init__(daemon=True)
        self._store = store
        self._running = threading.Event()
        self._sc_reader = SimConnectReader()
        self._fs_reader = FsuipcReader()
        self._start_time = time.time()

    def start_collecting(self):
        self._running.set()
        if not self.is_alive():
            self.start()

    def stop_collecting(self):
        self._running.clear()

    def run(self):
        if self._sc_reader.available:
            self._sc_reader.connect()
        if self._fs_reader.available:
            self._fs_reader.connect()
        while True:
            if not self._running.is_set():
                time.sleep(0.2)
                continue
            data = None
            if self._sc_reader.available and self._sc_reader.connect():
                data = self._sc_reader.read()
            if not data and self._fs_reader.available and self._fs_reader.connect():
                data = self._fs_reader.read()
            if not data:
                self._store.update({"source": "unavailable"})
                time.sleep(1.0)
                continue
            self._store.update(data)
            time.sleep(1.0)

    def _mock_data(self):
        elapsed = time.time() - self._start_time
        return {
            "altitude_ft": 3500 + 800 * math.sin(elapsed / 8.0),
            "heading_deg": (elapsed * 6) % 360,
            "airspeed_kt": 120 + 25 * math.sin(elapsed / 5.0),
            "vertical_speed_fpm": 400 * math.sin(elapsed / 4.0),
            "latitude": 31.25 + 0.05 * math.sin(elapsed / 15.0),
            "longitude": 121.50 + 0.05 * math.cos(elapsed / 15.0),
            "pitch_deg": 2.5 * math.sin(elapsed / 6.0),
            "bank_deg": 10 * math.sin(elapsed / 7.0),
            "fuel_total_gal": 56 - (elapsed / 1200.0),
            "source": "mock",
        }


class FlightDataApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("720x700")
        self.root.configure(bg="#10121a")

        self.lang_data = self._load_language_data()
        self.lang_code = self._load_language_code()
        self.lang_strings = self.lang_data.get(
            self.lang_code, self.lang_data.get("zh-CN", {})
        )
        self.status_key = "idle"
        self.source_key = "idle"
        port = _load_port_value()
        self.current_url = f"http://127.0.0.1:{port}"
        self.current_ip = "--"

        self.store = DataStore()
        self.collector = DataCollector(self.store)
        self.server_thread = None
        self.server = None

        self.theme = self._theme_dark()
        self.theme_dialog = None
        self.theme_frame = None
        self.theme_title = None
        self.theme_hint = None
        self.theme_actions = None
        self.theme_entries = {}
        self.theme_label_widgets = []
        self.theme_pick_buttons = []

        self.lang_frame = None
        self.lang_label = None
        self.lang_combo = None
        self.lang_options = []
        self.lang_codes = []

        self._build_ui()
        self._apply_language()
        self._apply_theme(self.theme)
        self._schedule_ui_refresh()
        self._show_notice_if_needed()

    def _load_language_data(self):
        path = os.path.join(os.getcwd(), "gui-lang.json")
        try:
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except Exception:
            return {"zh-CN": {}}

    def _load_language_code(self):
        path = os.path.join(os.getcwd(), "language.txt")
        try:
            with open(path, "r", encoding="utf-8") as handle:
                value = handle.read().strip()
            return value or "zh-CN"
        except Exception:
            return "zh-CN"

    def _save_language_code(self, code):
        path = os.path.join(os.getcwd(), "language.txt")
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(code)
        except Exception:
            pass

    def _t(self, key, **kwargs):
        value = self.lang_strings.get(key)
        if value is None:
            value = self.lang_data.get("zh-CN", {}).get(key, key)
        try:
            return value.format(**kwargs)
        except Exception:
            return value

    def _apply_language(self):
        self.lang_strings = self.lang_data.get(
            self.lang_code, self.lang_data.get("zh-CN", {})
        )
        self.root.title(self._t("app_title"))
        self.header_label.config(text=self._t("app_header"))
        self.sub_label.config(text=self._t("app_subtitle"))

        if self.lang_label is not None:
            self.lang_label.config(text=self._t("language_label"))
        if self.lang_combo is not None:
            order = ["zh-CN", "zh-TW", "en"]
            self.lang_codes = [code for code in order if code in self.lang_data]
            for code in self.lang_data:
                if code not in self.lang_codes:
                    self.lang_codes.append(code)
            self.lang_options = [self.lang_data[code].get("name", code) for code in self.lang_codes]
            self.lang_combo["values"] = self.lang_options
            try:
                idx = self.lang_codes.index(self.lang_code)
            except ValueError:
                idx = 0
                self.lang_code = self.lang_codes[0] if self.lang_codes else "zh-CN"
            self.lang_combo.current(idx)

        self.start_button.config(text=self._t("btn_start"))
        self.stop_button.config(text=self._t("btn_stop"))
        self.export_button.config(text=self._t("btn_export"))
        self._set_link(self.current_url)
        self._set_ip(self.current_ip)
        self._set_status(self.status_key)
        self._set_source(self.source_key)

        for key, label in self.field_label_widgets:
            label.config(text=self._t(key))

        if self.theme_button is not None:
            self.theme_button.config(text=self._t("theme_title"))
        if self.theme_dialog is not None and self.theme_dialog.winfo_exists():
            self.theme_dialog.title(self._t("theme_dialog_title"))
        if self.theme_title is not None:
            self.theme_title.config(text=self._t("theme_title"))
        if self.theme_hint is not None:
            self.theme_hint.config(text=self._t("theme_hint"))

        for key, label in self.theme_label_widgets:
            label.config(text=self._t(key))
        for button in self.theme_pick_buttons:
            button.config(text=self._t("theme_pick"))
        if hasattr(self, "light_button"):
            self.light_button.config(text=self._t("theme_light"))
        if hasattr(self, "dark_button"):
            self.dark_button.config(text=self._t("theme_dark"))
        if hasattr(self, "apply_button"):
            self.apply_button.config(text=self._t("theme_apply"))

    def _on_language_change(self, _event=None):
        if not self.lang_combo:
            return
        idx = self.lang_combo.current()
        if idx < 0 or idx >= len(self.lang_codes):
            return
        self.lang_code = self.lang_codes[idx]
        self._save_language_code(self.lang_code)
        self._apply_language()

    def _set_status(self, key):
        self.status_key = key
        state = self._t(f"status_{key}")
        self.status_label.config(text=self._t("status_prefix", state=state))

    def _set_source(self, key):
        self.source_key = key
        source_text = self._t(f"source_{key}") if f"source_{key}" in self.lang_strings else key
        self.source_label.config(text=self._t("source_prefix", source=source_text))

    def _set_link(self, url):
        self.current_url = url
        self.link_label.config(text=self._t("label_link", url=url))

    def _set_ip(self, ip):
        self.current_ip = ip
        self.ip_label.config(text=self._t("label_ip", ip=ip))

    def _build_language_selector(self):
        self.lang_frame = tk.Frame(self.root)
        self.lang_frame.pack(pady=(0, 8))

        self.lang_label = tk.Label(self.lang_frame, font=("Palatino", 10))
        self.lang_label.grid(row=0, column=0, padx=(0, 8))

        self.lang_combo = ttk.Combobox(self.lang_frame, state="readonly", width=14)
        self.lang_combo.grid(row=0, column=1)
        self.lang_combo.bind("<<ComboboxSelected>>", self._on_language_change)

    def _build_ui(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", font=("Helvetica", 12))

        self.header_label = tk.Label(
            self.root,
            text=self._t("app_header"),
            font=("Palatino", 20, "bold"),
        )
        self.header_label.pack(pady=(20, 10))

        self.sub_label = tk.Label(
            self.root,
            text=self._t("app_subtitle"),
            font=("Palatino", 12),
        )
        self.sub_label.pack(pady=(0, 8))

        self._build_language_selector()


        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10)

        self.start_button = ttk.Button(
            self.button_frame, text=self._t("btn_start"), command=self._start_collecting
        )
        self.start_button.grid(row=0, column=0, padx=8)

        self.stop_button = ttk.Button(
            self.button_frame, text=self._t("btn_stop"), command=self._stop_collecting
        )
        self.stop_button.grid(row=0, column=1, padx=8)

        self.export_button = ttk.Button(
            self.button_frame, text=self._t("btn_export"), command=self._start_server
        )
        self.export_button.grid(row=0, column=2, padx=8)

        self.link_label = tk.Label(
            self.root,
            text=self._t("label_link", url=self.current_url),
            font=("Palatino", 11),
        )
        self.link_label.pack(pady=(10, 6))

        self.ip_label = tk.Label(
            self.root,
            text=self._t("label_ip", ip=self.current_ip),
            font=("Palatino", 10),
        )
        self.ip_label.pack(pady=(0, 10))

        self.card_frame = tk.Frame(self.root, padx=20, pady=20)
        self.card_frame.pack(fill="both", expand=True, padx=28, pady=(0, 20))

        self.value_labels = {}
        self.card_frames = []
        self.card_name_labels = []
        self.field_label_widgets = []
        fields = [
            ("field_altitude", "altitude_ft", "ft"),
            ("field_heading", "heading_deg", "deg"),
            ("field_speed", "airspeed_kt", "kt"),
            ("field_vs", "vertical_speed_fpm", "fpm"),
            ("field_lat", "latitude", ""),
            ("field_lon", "longitude", ""),
            ("field_pitch", "pitch_deg", "deg"),
            ("field_bank", "bank_deg", "deg"),
            ("field_fuel", "fuel_total_gal", "gal"),
        ]

        for idx, (label_key, key, unit) in enumerate(fields):
            row = idx // 2
            col = idx % 2
            frame = tk.Frame(self.card_frame)
            frame.grid(row=row, column=col, sticky="w", padx=8, pady=8)
            self.card_frames.append(frame)

            name = tk.Label(frame, text=self._t(label_key), font=("Palatino", 12))
            name.pack(anchor="w")
            self.card_name_labels.append(name)
            self.field_label_widgets.append((label_key, name))

            value = tk.Label(frame, text="--", font=("Courier New", 20, "bold"))
            value.pack(anchor="w")
            self.value_labels[key] = (value, unit)

        self.status_label = tk.Label(
            self.root,
            text="",
            font=("Palatino", 11),
        )
        self.status_label.pack(pady=(0, 4))

        self.source_label = tk.Label(
            self.root,
            text="",
            font=("Palatino", 10),
        )
        self.source_label.pack(pady=(0, 10))

        self._build_theme_controls()

    def _build_theme_controls(self):
        self.theme_button = ttk.Button(
            self.root, text=self._t("theme_button"), command=self._open_theme_dialog
        )
        self.theme_button.pack(pady=(0, 16))

    def _open_theme_dialog(self):
        if self.theme_dialog and self.theme_dialog.winfo_exists():
            self.theme_dialog.lift()
            return
        dialog = tk.Toplevel(self.root)
        dialog.title(self._t("theme_dialog_title"))
        dialog.configure(bg=self.theme["bg"])
        dialog.geometry("420x420")
        dialog.transient(self.root)
        dialog.grab_set()
        self.theme_dialog = dialog

        container = tk.Frame(dialog, bg=self.theme["bg"])
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=self.theme["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=self.theme["bg"])
        canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", on_configure)

        def on_mousewheel(event):
            if event.delta:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        def on_close():
            canvas.unbind_all("<MouseWheel>")
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", on_close)

        self._build_theme_panel(inner)
        self._apply_theme(self.theme)

    def _build_theme_panel(self, parent):
        self.theme_frame = parent
        self.theme_label_widgets = []
        self.theme_pick_buttons = []
        self.theme_title = tk.Label(
            self.theme_frame,
            text=self._t("theme_title"),
            font=("Palatino", 12, "bold"),
        )
        self.theme_title.grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        self.theme_entries = {}
        fields = [
            ("theme_bg", "bg"),
            ("theme_panel", "panel"),
            ("theme_text", "text"),
            ("theme_muted", "muted"),
            ("theme_accent", "accent"),
            ("theme_border", "border"),
        ]

        for idx, (label, key) in enumerate(fields, start=1):
            name = tk.Label(self.theme_frame, text=self._t(label), font=("Palatino", 10))
            name.grid(row=idx, column=0, sticky="w", pady=4)

            entry = tk.Entry(self.theme_frame, width=12)
            entry.grid(row=idx, column=1, padx=8, pady=4, sticky="w")
            entry.insert(0, self.theme.get(key, "#000000"))
            self.theme_entries[key] = entry
            self.theme_label_widgets.append((label, name))

            button = ttk.Button(
                self.theme_frame,
                text=self._t("theme_pick"),
                command=lambda k=key: self._pick_color(k),
            )
            button.grid(row=idx, column=2, padx=4, pady=4, sticky="w")
            self.theme_pick_buttons.append(button)

        self.theme_hint = tk.Label(
            self.theme_frame,
            text=self._t("theme_hint"),
            font=("Palatino", 9),
        )
        self.theme_hint.grid(row=7, column=0, columnspan=3, sticky="w", pady=(6, 8))

        self.theme_actions = tk.Frame(self.theme_frame)
        self.theme_actions.grid(row=8, column=0, columnspan=3, sticky="w")

        self.light_button = ttk.Button(
            self.theme_actions, text=self._t("theme_light"), command=self._apply_light_theme
        )
        self.light_button.grid(row=0, column=0, padx=(0, 8))

        self.dark_button = ttk.Button(
            self.theme_actions, text=self._t("theme_dark"), command=self._apply_dark_theme
        )
        self.dark_button.grid(row=0, column=1, padx=(0, 8))

        self.apply_button = ttk.Button(
            self.theme_actions, text=self._t("theme_apply"), command=self._apply_custom_theme
        )
        self.apply_button.grid(row=0, column=2)

    def _pick_color(self, key):
        current = self.theme_entries[key].get().strip() or "#000000"
        color = colorchooser.askcolor(color=current)[1]
        if color:
            self.theme_entries[key].delete(0, tk.END)
            self.theme_entries[key].insert(0, color)

    def _apply_custom_theme(self):
        theme = {}
        for key, entry in self.theme_entries.items():
            value = entry.get().strip()
            if not self._is_hex_color(value):
                messagebox.showerror(self._t("color_error_title"), self._t("color_error_body", value=value))
                return
            theme[key] = value
        self.theme = theme
        self._apply_theme(theme)

    def _apply_light_theme(self):
        self.theme = self._theme_light()
        self._sync_theme_entries(self.theme)
        self._apply_theme(self.theme)

    def _apply_dark_theme(self):
        self.theme = self._theme_dark()
        self.theme_dialog = None
        self.theme_frame = None
        self.theme_title = None
        self.theme_hint = None
        self.theme_actions = None
        self.theme_entries = {}
        self._sync_theme_entries(self.theme)
        self._apply_theme(self.theme)

    def _sync_theme_entries(self, theme):
        for key, entry in self.theme_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, theme.get(key, "#000000"))

    def _apply_theme(self, theme):
        self.root.configure(bg=theme["bg"])
        self.header_label.configure(bg=theme["bg"], fg=theme["text"])
        self.sub_label.configure(bg=theme["bg"], fg=theme["muted"])
        if self.lang_frame is not None:
            self.lang_frame.configure(bg=theme["bg"])
        if self.lang_label is not None:
            self.lang_label.configure(bg=theme["bg"], fg=theme["text"])
        self.button_frame.configure(bg=theme["bg"])
        self.link_label.configure(bg=theme["bg"], fg=theme["accent"])
        self.ip_label.configure(bg=theme["bg"], fg=theme["muted"])
        self.status_label.configure(bg=theme["bg"], fg=theme["text"])
        self.source_label.configure(bg=theme["bg"], fg=theme["muted"])

        self.card_frame.configure(bg=theme["bg"])
        for frame in self.card_frames:
            frame.configure(bg=theme["panel"])
        for label in self.card_name_labels:
            label.configure(bg=theme["panel"], fg=theme["muted"])
        for label, _unit in self.value_labels.values():
            label.configure(bg=theme["panel"], fg=theme["text"])

        if self.theme_frame is not None:
            self.theme_frame.configure(bg=theme["bg"])
        if self.theme_title is not None:
            self.theme_title.configure(bg=theme["bg"], fg=theme["text"])
        if self.theme_hint is not None:
            self.theme_hint.configure(bg=theme["bg"], fg=theme["muted"])
        if self.theme_actions is not None:
            self.theme_actions.configure(bg=theme["bg"])
        if self.theme_frame is not None:
            for child in self.theme_frame.winfo_children():
                if isinstance(child, tk.Label):
                    child.configure(bg=theme["bg"], fg=theme["text"])

        for entry in self.theme_entries.values():
            entry.configure(
                bg=theme["panel"],
                fg=theme["text"],
                insertbackground=theme["text"],
                highlightbackground=theme["border"],
                highlightcolor=theme["accent"],
            )

        self.style.configure(
            "TButton",
            background=theme["panel"],
            foreground=theme["text"],
            bordercolor=theme["border"],
        )
        self.style.map(
            "TButton",
            background=[("active", theme["accent"]), ("pressed", theme["accent"])],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )

    def _theme_light(self):
        return {
            "bg": "#f5f7fa",
            "panel": "#ffffff",
            "text": "#0f172a",
            "muted": "#475569",
            "accent": "#1d4ed8",
            "border": "#d7dde7",
        }

    def _theme_dark(self):
        return {
            "bg": "#10121a",
            "panel": "#1f2432",
            "text": "#f4f1de",
            "muted": "#c7cbe0",
            "accent": "#9ad1d4",
            "border": "#2b3246",
        }

    def _is_hex_color(self, value):
        if not value.startswith("#") or len(value) != 7:
            return False
        for ch in value[1:]:
            if ch not in "0123456789abcdefABCDEF":
                return False
        return True

    def _start_collecting(self):
        self.collector.start_collecting()
        self._set_status("running")

    def _stop_collecting(self):
        self.collector.stop_collecting()
        self._set_status("stopped")

    def _start_server(self):
        if self.server_thread and self.server_thread.is_alive():
            return
        ip = self._get_local_ip()
        port = self._load_port()
        self.server = DataServer(
            "0.0.0.0", port, self.store.snapshot, static_dir=os.getcwd()
        )
        self.server_thread = threading.Thread(target=self.server.start, daemon=True)
        self.server_thread.start()
        if ip:
            self._set_link(f"http://{ip}:{port}")
            self._set_ip(ip)
        self._set_status("exported")

    def _schedule_ui_refresh(self):
        data = self.store.snapshot()
        for key, (label, unit) in self.value_labels.items():
            value = data.get(key, 0.0)
            if isinstance(value, float):
                text = f"{value:.2f} {unit}".strip()
            else:
                text = f"{value} {unit}".strip()
            label.config(text=text)
        source = data.get("source", "idle")
        self._set_source(source)
        self.root.after(800, self._schedule_ui_refresh)

    def _get_local_ip(self):
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.connect(("8.8.8.8", 80))
            ip = sock.getsockname()[0]
        except OSError:
            ip = None
        finally:
            try:
                sock.close()
            except Exception:
                pass
        return ip

    def _load_port(self):
        return _load_port_value()

    def _show_notice_if_needed(self):
        flag_path = os.path.join(os.getcwd(), "notice_flag.txt")
        try:
            with open(flag_path, "r", encoding="utf-8") as handle:
                if handle.read().strip() == "1":
                    return
        except Exception:
            pass
        self._open_notice_dialog(flag_path)

    def _open_notice_dialog(self, flag_path):
        dialog = tk.Toplevel(self.root)
        dialog.title(self._t("notice_title"))
        dialog.configure(bg=self.theme["bg"])
        dialog.geometry("460x340")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        title = tk.Label(
            dialog,
            text=self._t("notice_title"),
            fg=self.theme["text"],
            bg=self.theme["bg"],
            font=("Palatino", 14, "bold"),
        )
        title.pack(pady=(18, 10))

        content = (
            self._t("notice_line1") + "\n"
            + self._t("notice_line2") + "\n"
            + self._t("notice_line3") + "\n"
            + self._t("notice_line4") + "\n"
            + self._t("notice_line5") + "\n"
            + self._t("notice_line6") + "\n"
        )
        body = tk.Label(
            dialog,
            text=content,
            fg=self.theme["muted"],
            bg=self.theme["bg"],
            font=("Palatino", 11),
            justify="left",
        )
        body.pack(padx=20, pady=(0, 12), anchor="w")

        self._dont_show_var = tk.BooleanVar(value=False)
        checkbox = ttk.Checkbutton(
            dialog,
            text=self._t("notice_checkbox"),
            variable=self._dont_show_var,
        )
        checkbox.pack(pady=(0, 14))

        def close_dialog():
            if self._dont_show_var.get():
                try:
                    with open(flag_path, "w", encoding="utf-8") as handle:
                        handle.write("1")
                except Exception:
                    pass
            dialog.destroy()

        confirm = ttk.Button(dialog, text=self._t("notice_confirm"), command=close_dialog)
        confirm.pack(pady=(0, 16))


def main():
    port = _load_port_value()
    _ensure_firewall_rule(port)
    root = tk.Tk()
    app = FlightDataApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
