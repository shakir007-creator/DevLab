VERSION = "2.0.0"  # Semantic Versioning: Update this for every new release

import pandas as pd
import re
import os
import tkinter as tk
from tkinter import Text, END, filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from datetime import datetime
import ipaddress
import sys
import threading
import queue
import time
import requests
import ctypes
import json

# Update from
GITHUB_RELEASE_API = ("https://api.github.com/repos/PacketForge007/SonicWallTool/releases/latest")

DEFAULT_PASSWORD = "Admin@12345"

def version_tuple(v):
    try:
        return tuple(map(int, v.strip().split(".")))
    except Exception:
        return (0,)

def copy_with_progress(src, dst, progress_window):
    total = os.path.getsize(src)
    copied = 0
    buf_size = 1024 * 1024  # 1MB
    start_time = time.time()

    with open(src, "rb") as fsrc, open(dst, "wb") as fdst:
        while True:
            if progress_window.cancelled:
                raise Exception("Update cancelled by user")

            buf = fsrc.read(buf_size)
            if not buf:
                break

            fdst.write(buf)
            copied += len(buf)

            elapsed = time.time() - start_time
            speed = (copied / 1024 / 1024) / elapsed if elapsed > 0 else 0
            percent = int((copied / total) * 100)
            remaining = total - copied
            eta = int((remaining / 1024 / 1024) / speed) if speed > 0 else "--"

            progress_window.update_download(percent, speed, eta)


def download_with_progress(url, dst, progress_window):
    response = requests.get(url, stream=True, timeout=30)
    response.raise_for_status()

    total = int(response.headers.get("content-length", 0))
    copied = 0
    start_time = time.time()

    with open(dst, "wb") as f:
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if progress_window.cancelled:
                raise Exception("Update cancelled by user")

            if chunk:
                f.write(chunk)
                copied += len(chunk)

                elapsed = time.time() - start_time
                speed = (copied / 1024 / 1024) / elapsed if elapsed > 0 else 0
                percent = int((copied / total) * 100) if total else 0
                remaining = total - copied
                eta = int((remaining / 1024 / 1024) / speed) if speed > 0 else "--"

                progress_window.update_download(percent, speed, eta)


def launch_updater(latest_version, download_url, root):
    base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    updater_path = os.path.join(base_path, "SonicUpdater.exe")

    if not os.path.exists(updater_path):
        messagebox.showerror("Updater Missing", "Updater program not found.")
        return False

    try:
        response = requests.head(download_url, timeout=10)
        response.raise_for_status()
    except Exception:
        messagebox.showerror(
            "Update Not Found",
            f"Update file not found on server:\n{download_url}"
        )
        return False

    try:
        params = f'"{latest_version}" "{download_url}"'

        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            os.path.abspath(updater_path),
            params,
            os.path.dirname(os.path.abspath(updater_path)),
            1
        )

        if result <= 32:
            messagebox.showwarning(
                "Update Cancelled",
                "Update was cancelled or administrator access was denied."
            )
            return False

        root.destroy()
        return True

    except Exception as e:
        messagebox.showerror("Update Failed", str(e))
        return False


def check_for_updates(auto=False):
    try:
        response = requests.get(GITHUB_RELEASE_API, timeout=10)
        response.raise_for_status()

        release = response.json()

        latest_version = release["tag_name"].lstrip("v")
        download_url = None

        for asset in release["assets"]:
            name = asset["name"]

            if (
                name.startswith("SonicWallTool_v")
                and name.endswith(".exe")
                and "Installer" not in name
                and "Updater" not in name
            ):
                download_url = asset["browser_download_url"]
                break

        if not download_url:
            raise Exception("Application executable not found in GitHub release.")

    except Exception as e:
        if not auto:
            messagebox.showerror(
                "Update Check Failed",
                str(e)
            )
        return

    except Exception:
        if not auto:
            messagebox.showerror(
                "Update Check Failed",
                "Could not connect to GitHub."
            )
        return

    current_version = version_tuple(VERSION)
    new_version = version_tuple(latest_version)

    if new_version > current_version:
        answer = True if auto else messagebox.askyesno(
            "Update Available",
            f"Version {latest_version} is available.\nUpdate now?"
        )

        if answer:
            file_menu.entryconfig(EXIT_MENU_INDEX, state="disabled")

            started = launch_updater(latest_version, download_url, root)

            if not started:
                file_menu.entryconfig(EXIT_MENU_INDEX, state="normal")

    else:
        if not auto:
            messagebox.showinfo(
                "No Update",
                "You already have the latest version."
            )


# User template
USER_TEMPLATE_COLUMNS = ["Username", "Password", "Group"]
USER_TEMPLATE_DATA = [
    ["User1", "Admin@1234", ""],
    ["User2", "Admin@1234", "Group1"],
    ["User3", "Admin@1234", "Group1, Group2"]
]

# Address Object template
ADDRESS_TEMPLATE_COLUMNS = ["Name", "Address", "Zone", "Group"]
ADDRESS_TEMPLATE_DATA = [
    ["Abc", "172.16.1.0/24", "LAN", "Group1"],
    ["Def", "192.168.10.1", "LAN", "Group1, Group2"],
    ["Ghi", "100.100.100.100", "WAN", "Telnet Grp"],
    ["Jkl_FQDN", "google.com", "WAN", "FQDN_Grp1"],
    ["Mno", "*.facebook.com", "WAN", "FQDN_Grp2"],
    ["Pqr", "00:11:22:aa:bb:cc", "LAN", "Group2"],
    ["Stu", "10.10.1.1-10.10.1.10", "LAN", ""]
]

# Service Object template
SERVICE_TEMPLATE_COLUMNS = ["Name", "Protocol", "Port", "Group"]
SERVICE_TEMPLATE_DATA = [
    ["Abc", "tcp", 22, "SSH_Grp"],
    ["Def", "udp", 22, "SSH_Grp, SSH_Grp2"],
    ["Ghi", "tcp", 23, "Telnet_Grp"],
    ["Yz", "tcp", "25-26", ""]
]

# Group template
GROUP_TEMPLATE_COLUMNS = ["Group", "Member"]
GROUP_TEMPLATE_DATA = [
    ["Group1", ""],
    ["Group2", "User1"],
    ["Group3", "User2, User3"]
]

# DHCP template
DHCP_TEMPLATE_COLUMNS = ["Name", "MAC", "IP", "Mask", "GW", "DNS"]
DHCP_TEMPLATE_DATA = [
    ["PC-01", "00:11:22:33:44:55", "192.168.1.10", "255.255.255.0", "192.168.1.1", "8.8.8.8"],
    ["PC-02", "AA:BB:CC:DD:EE:FF", "192.168.1.11", "255.255.255.0", "192.168.1.1", "8.8.8.8, 8.8.4.4"]
]


def export_template(columns, data, default_filename):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=default_filename
    )
    if file_path:
        try:
            df = pd.DataFrame(data, columns=columns)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Template Exported", f"Template saved as:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


def is_valid_password(pwd):
    return (
        len(pwd) >= 8 and
        re.search(r'[A-Z]', pwd) and
        re.search(r'[a-z]', pwd) and
        re.search(r'[0-9]', pwd) and
        re.search(r'[^A-Za-z0-9]', pwd)
    )

def needs_quotes(s):
    return bool(re.search(r'[\s_\-\W]', s))

def generate_user_command(username, password, group_list):
    logs = []

    if not password or pd.isna(password):
        password = DEFAULT_PASSWORD

    quoted_username = f'"{username}"' if needs_quotes(username) else username
    quoted_password = f'"{password}"'
    cmd = f"user {quoted_username} password {quoted_password}"
    return [cmd], logs

# User Full Script
def generate_full_script(user_rows):
    all_cmds = []
    logs = []
    group_map = {}

    for index, row in user_rows.iterrows():

        row_number = index + 2  # +2 to account for header and 0-based index

        username = "" if pd.isna(row.get("Username")) else str(row.get("Username")).strip()
        password = "" if pd.isna(row.get("Password")) else str(row.get("Password")).strip()
        raw_groups = "" if pd.isna(row.get("Group")) else str(row.get("Group")).strip()

        if not username:
            logs.append(f"[Row {row_number}] Skipped: Empty username")
            continue

        # ✅ Parse "Grp1, Grp2"
        if isinstance(raw_groups, str):
            group_list = [g.strip() for g in raw_groups.split(",") if g.strip()]
        else:
            group_list = []

        cmds, log = generate_user_command(username, password, group_list)
        logs.extend(log)

        if cmds:
            all_cmds.extend(cmds)

            for grp in group_list:
                group_map.setdefault(grp, []).append(username)

    if all_cmds:
        all_cmds.append("commit")
        all_cmds.append("\n")

    for grp, users in group_map.items():
        unique_users = list(dict.fromkeys(users))
        quoted_grp = f'"{grp}"' if needs_quotes(grp) else grp
        all_cmds.append(f"group {quoted_grp}")
        for user in unique_users:
            quoted_user = f'"{user}"' if needs_quotes(user) else user
            all_cmds.append(f"member {quoted_user}")
        all_cmds.append("exit")

    if group_map:
        all_cmds.append("commit")

    return all_cmds, logs


# Detect Address Type
def detect_type(address):
    try:
        if '/' in address:
            ip_part, mask_part = address.split("/", 1)
            if (re.match(r'^\d+\.\d+\.\d+\.\d+$', mask_part) and mask_part == "255.255.255.255") or mask_part == "32":
                return 'host'
            ipaddress.IPv4Network(address, strict=False)
            return 'network'
        ipaddress.IPv4Address(address)
        return 'host'
    except:
        pass
    if re.match(r'^\d+\.\d+\.\d+\.\d+\s*-\s*\d+\.\d+\.\d+\.\d+$', address.strip()):
        return 'range'
    if re.match(r'^(\*\.)?([A-Za-z0-9-]+\.)+[A-Za-z]{2,}$', address.strip()):
        return 'fqdn'
    if re.match(r'^([0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}$', address.strip()):
        return 'mac'
    return 'unknown'

def generate_address_command(name, address, zone):
    if not address or pd.isna(address) or str(address).strip().lower() == "nan":
        return [], [f"# Skipped {name or 'Unnamed'}: Address is missing"]
    
    if not zone or pd.isna(zone) or str(zone).strip().lower() == "nan":
        return [], [f"# Skipped {name or address or 'Unnamed'}: Zone is missing"]

    name = name if pd.notna(name) and str(name).strip().lower() != "nan" and str(name).strip() else address
    quoted_name = f'"{name}"' if needs_quotes(name) else name
    quoted_zone = f'"{zone}"'
    
    addr_type = detect_type(address)

    if addr_type == 'host':
        if '/' in address:
            ip_part, mask_part = address.split("/", 1)
            ip_part.strip()
            return [f"address-object ipv4 {quoted_name} host {ip_part} zone {quoted_zone}"], []
        return [f"address-object ipv4 {quoted_name} host {address} zone {quoted_zone}"], []
    elif addr_type == 'network':
        net = ipaddress.IPv4Network(address, strict=False)
        return [f"address-object ipv4 {quoted_name} network {net.network_address} /{net.prefixlen} zone {quoted_zone}"], []
    elif addr_type == 'range':
        ip1, ip2 = [ip.strip() for ip in address.split('-')]
        return [f"address-object ipv4 {quoted_name} range {ip1} {ip2} zone {quoted_zone}"], []
    elif addr_type == 'fqdn':
        return [f"address-object fqdn {quoted_name} domain {address.strip()} zone {quoted_zone}", "exit"], []
    elif addr_type == 'mac':
        return [f"address-object mac {quoted_name} address {address.strip()} zone {quoted_zone}", "exit"], []
    else:
        return [], [f"# Skipped {name}: Unknown address format '{address}'"]

# New function to generate service-object command
def generate_service_command(name, protocol, port):
    logs = []
    if not name or pd.isna(name) or str(name).strip().lower() == "nan" or not str(name).strip():
        logs.append("# Skipped: Empty service name")
        return [], logs
    if not protocol or pd.isna(protocol) or str(protocol).strip().lower() == "nan" or not str(protocol).strip():
        logs.append(f"# Skipped {name}: Protocol missing")
        return [], logs
    if not port or pd.isna(port) or str(port).strip().lower() == "nan" or not str(port).strip():
        logs.append(f"# Skipped {name}: Port missing")
        return [], logs

    name = str(name).strip()
    protocol = str(protocol).strip().lower()
    port = str(port).strip()

    valid_protocols = {
        "6over4", "ah", "eigrp", "esp", "gre", "icmp", "icmpv6", "icmpv6-custom",
        "igmp", "l2tp", "ospf", "pim", "tcp", "udp"
    }
    if protocol not in valid_protocols:
        logs.append(f"# Skipped {name}: Invalid protocol '{protocol}'")
        return [], logs

    # Check port format
    if '-' in port:
        parts = port.split('-')
        if len(parts) != 2:
            logs.append(f"# Skipped {name}: Invalid port range '{port}'")
            return [], logs
        p_start, p_end = parts[0].strip(), parts[1].strip()
        if not p_start.isdigit() or not p_end.isdigit():
            logs.append(f"# Skipped {name}: Port range contains non-numeric values '{port}'")
            return [], logs
        quoted_names= f'"{name}"' if needs_quotes(name) else name
        cmd = f'service-object {quoted_names} {protocol} {p_start} {p_end}'
    else:
        if not port.isdigit():
            logs.append(f"# Skipped {name}: Invalid port '{port}'")
            return [], logs

        port_num = int(port)

        if not (1 <= port_num <= 65535):
            logs.append(f"# Skipped {name}: Port out of range '{port}'")
            return [], logs

        quoted_names= f'"{name}"' if needs_quotes(name) else name
        cmd = f'service-object {quoted_names} {protocol} {port} {port}'

    return [cmd], logs

def generate_group_command(group_name, members, row_number=None):
    logs = []

    def quote_name(name):
        """
        Wrap name in double quotes if it contains spaces, hyphens, underscores, or other non-alphanumeric characters.
        """
        if not name:
            return ""
        name = str(name).strip()
        if any(c in name for c in " -_") or not name.isalnum():
            return f'"{name}"'
        return name

    if not group_name or pd.isna(group_name) or not str(group_name).strip() or str(group_name).strip().lower() == "nan":
        row_info = f"row {row_number}" if row_number is not None else ""
        logs.append(f"# Skipped {row_info}: Empty group name")
        return [], logs

    group_name = quote_name(group_name)
    cmds = [f"group {group_name}"]

    if members and pd.notna(members):
        member_list = [m.strip() for m in str(members).split(',') if m.strip()]
        for m in member_list:
            m_quoted = quote_name(m)
            cmds.append(f"member {m_quoted}")

    cmds.append("exit")
    return cmds, logs

# DHCP Static Binding
def normalize_mac(mac):
    if pd.isna(mac):
        return ""
    mac = str(mac).lower()
    mac = re.sub(r'[^a-f0-9]', '', mac)
    if len(mac) != 12:
        return ""
    return ":".join(mac[i:i+2] for i in range(0, 12, 2))


def parse_dns(dns_str):
    return [ip.strip() for ip in str(dns_str).split(",") if ip.strip()]


def format_name(name):
    if re.search(r"[^A-Za-z0-9]", str(name)):
        return f'"{name}"'
    return str(name)


def generate_dhcp_cli_script(df, log_lines):
    cli_lines = ["dhcp-server\n"]
    required_cols = ["Name", "MAC", "IP", "Mask", "GW", "DNS"]

    for index, row in df.iterrows():
        missing = [
            c for c in required_cols
            if c not in row or pd.isna(row[c]) or str(row[c]).strip() == ""
        ]
        if missing:
            log_lines.append(
                f"[Row {index+2}] Skipped: Missing {', '.join(missing)}"
            )
            continue

        mac = normalize_mac(row["MAC"])
        if not mac:
            log_lines.append(
                f"[Row {index+2}] Skipped: Invalid MAC address"
            )
            continue

        cli_lines.extend([
            f"scope static {row['IP']} {mac}",
            f"name {format_name(row['Name'])}",
            f"netmask {row['Mask']}",
            f"default-gateway {row['GW']}",
        ])

        dns = parse_dns(row["DNS"])
        if len(dns) > 0:
            cli_lines.append(f"dns server static primary {dns[0]}")
        if len(dns) > 1:
            cli_lines.append(f"dns server static secondary {dns[1]}")
        if len(dns) > 2:
            cli_lines.append(f"dns server static tertiary {dns[2]}")

        cli_lines.append("enable\nexit\n")

    return "\n".join(cli_lines)


class SonicToolApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()      # Hide window while building GUI
        root.title("SonicWallTool")
        root.resizable(True, True)

        # Get the folder where the script or exe is located
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        self.settings_file = os.path.join(base_path, "window_settings.json")

        self.user_files = []
        self.addr_files = []
        self.svc_files = []
        self.grp_files = []
        self.dhcp_files = []

        self.stop_requested = False

        self.tree_state = {}

        self.log_search_state = {}

        # Build the path to your icon dynamically
        icon_path = os.path.join(base_path, "Sonic.ico")  # Icon must be in the same folder

        # root.iconbitmap(icon_path)
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
        style = ttk.Style()
        style.theme_use("default")
        self.tabs = ttk.Notebook(root)
        self.tab1 = ttk.Frame(self.tabs)
        self.tab2 = ttk.Frame(self.tabs)
        self.tab3 = ttk.Frame(self.tabs)
        self.tab4 = ttk.Frame(self.tabs)
        self.tab5 = ttk.Frame(self.tabs)
        self.tabs.add(self.tab1, text="Local User Generator")
        self.tabs.add(self.tab2, text="Address Object Generator")
        self.tabs.add(self.tab3, text="Service Object Generator")
        self.tabs.add(self.tab4, text="User Group Generator")
        self.tabs.add(self.tab5, text="DHCP Static Generator")
        self.tabs.pack(expand=1, fill="both")
        self.build_user_tab(self.tab1)
        self.build_address_tab(self.tab2)
        self.build_service_tab(self.tab3)
        self.build_group_tab(self.tab4)
        self.build_dhcp_tab(self.tab5)

        self.gui_queue = queue.Queue()
        self.process_gui_queue()

        # Let Tkinter calculate widget sizes
        self.root.update_idletasks()

        # Load saved size or create default size
        self.load_window_settings()

        # Recalculate everything after geometry change
        self.root.update_idletasks()

        # Finally show the window
        self.root.deiconify()

        # Handle window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.bind_all("<Control-f>", self.focus_log_search)

    # Bind keyboard shortcuts for Treeview
    def bind_treeview_shortcuts(self, tree, file_list, count_label):
        tree.bind("<ButtonRelease-1>", lambda e: self.on_tree_click_release(tree, e), add="+")
        tree.bind("<Control-Button-1>", lambda e: self.on_tree_ctrl_click(tree, e))

        tree.bind("<Up>", lambda e: self.tree_up(tree))
        tree.bind("<Down>", lambda e: self.tree_down(tree))

        tree.bind("<Home>", lambda e: self.tree_home(tree))
        tree.bind("<End>", lambda e: self.tree_end(tree))
        tree.bind("<Delete>", lambda e: self.remove_selected_file(tree, file_list, count_label))

        tree.bind("<Control-a>", lambda e: self.tree_select_all(tree))

        tree.bind("<Shift-Home>", lambda e: self.tree_shift_home(tree))
        tree.bind("<Shift-End>", lambda e: self.tree_shift_end(tree))

        tree.bind("<Shift-Up>", lambda e: self.tree_shift_up(tree))
        tree.bind("<Shift-Down>", lambda e: self.tree_shift_down(tree))


    def center_window(self, width, height):

        self.root.update_idletasks()  # Update "requested size" from geometry manager

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2, 0)

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def load_window_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r") as f:
                    settings = json.load(f)

                geometry = settings.get("geometry")

                if geometry:
                    self.root.geometry(geometry)
                    self.root.update_idletasks()
                    return

            except Exception:
                pass

        # Default size = 80% of screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        width = int(screen_width * 0.80)
        height = int(screen_height * 0.80)

        self.center_window(width, height)

    def save_window_settings(self):
        if self.root.state() == "normal":
            settings = {
                "geometry": self.root.geometry()
            }

            try:
                with open(self.settings_file, "w") as f:
                    json.dump(settings, f)
            except Exception:
                pass

    def on_close(self):
        self.save_window_settings()
        self.root.destroy()

    def build_common_layout(self, tab, title1):
        ttk.Label(tab, text=title1).grid(row=0, column=0, sticky="w", padx=5, pady=5)
 
        # ===== Treeview Frame =====
        tree_frame = ttk.Frame(tab)
        tree_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        columns = ("filename", "path")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=6, selectmode="extended")
        tree.heading("filename", text="File Name")
        tree.heading("path", text="Full Path")
        tree.column("filename", width=180, anchor="w")
        tree.column("path", width=650, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        # Vertical Scrollbar
        tree_vscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree_vscroll.grid(row=0, column=1, sticky="ns")
        # Horizontal Scrollbar
        tree_hscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree_hscroll.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=tree_vscroll.set, xscrollcommand=tree_hscroll.set)

        # Allow Treeview to resize
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # ===== Output Folder =====
        ttk.Label(tab, text="Select output folder:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        entry_output = ttk.Entry(tab, width=60)
        entry_output.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        btn_output = ttk.Button(tab, text="📁 Browse", command=lambda: self.browse(entry_output))
        btn_output.grid(row=2, column=3, padx=5)

        # ===== Logs =====
        log_frame = ttk.Frame(tab)
        log_frame.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky="nsew")
        log_header = ttk.Frame(tab)
        log_header.grid(row=3, column=0, columnspan=4, sticky="ew", padx=5, pady=(5, 0))
        ttk.Label(log_header, text="Logs:").pack(side="left")
        ttk.Frame(log_header).pack(side="left", expand=True, fill="x")
        ttk.Label(log_header, text="Search in Logs:").pack(side="left")
        txt_log = Text(log_frame, height=18, wrap="none")
        txt_log.tag_configure("search_match", background="yellow", foreground="black")
        txt_log.tag_configure("current_match", background="orange", foreground="black")
        txt_log.grid(row=0, column=0, sticky="nsew")
        log_search_var = tk.StringVar()
        log_search_entry = ttk.Entry(log_header, textvariable=log_search_var, width=25)
        log_search_entry.pack(side="left", padx=(5, 5))
        log_match_label = ttk.Label(log_header, text="0 / 0")
        log_match_label.pack(side="left", padx=(8, 0))
        log_prev_btn = ttk.Button(log_header, text="Previous", width=10, command=lambda txt=txt_log, label=log_match_label: self.goto_previous_match(txt, label))
        log_prev_btn.pack(side="left", padx=(0, 2))
        log_next_btn = ttk.Button(log_header, text="Next", width=8, command=lambda txt=txt_log, label=log_match_label: self.goto_next_match(txt, label))
        log_next_btn.pack(side="left")
        log_search_var.trace_add("write", lambda *args, txt=txt_log, var=log_search_var, label=log_match_label: self.highlight_log_matches(txt, var, label))
        # Keyboard shortcuts
        log_search_entry.bind("<Return>", lambda e, txt=txt_log, label=log_match_label: (self.goto_next_match(txt, label), "break")[-1])
        log_search_entry.bind("<Shift-Return>", lambda e, txt=txt_log, label=log_match_label: (self.goto_previous_match(txt, label), "break")[-1])
        log_search_entry.bind("<F3>", lambda e, txt=txt_log, label=log_match_label: (self.goto_next_match(txt, label), "break")[-1])
        log_search_entry.bind("<Shift-F3>", lambda e, txt=txt_log, label=log_match_label: (self.goto_previous_match(txt, label), "break")[-1])
        # Vertical Scrollbar
        log_vscroll = ttk.Scrollbar(log_frame, orient="vertical", command=txt_log.yview)
        log_vscroll.grid(row=0, column=1, sticky="ns")
        # Horizontal Scrollbar
        log_hscroll = ttk.Scrollbar(log_frame, orient="horizontal", command=txt_log.xview)
        log_hscroll.grid(row=1, column=0, sticky="ew")
        txt_log.configure(yscrollcommand=log_vscroll.set, xscrollcommand=log_hscroll.set)
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        #==== Buttons =====
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=7, column=1, columnspan=2, pady=10)

        #==== Progress Bar =====
        progress = ttk.Progressbar(tab, orient="horizontal", mode="determinate", length=300)
        progress.grid(row=6, column=1, columnspan=2, sticky="ew", padx=5, pady=(0, 5))

        #==== Status Label =====
        ttk.Label(tab, text="Status:").grid(row=5, column=0, sticky="w", padx=5, pady=(0, 5))
        status = ttk.Label(tab, text="Ready")
        status.grid(row=5, column=1, sticky="w", padx=5, pady=(0, 5))

        tab.grid_columnconfigure(1, weight=1)

        tab.grid_rowconfigure(0, weight=1)   # Treeview
        tab.grid_rowconfigure(4, weight=1)   # Log box

        self.log_search_state[txt_log] = {"matches": [], "current": -1}

        if not hasattr(self, "log_search_entries"):
             self.log_search_entries = {}

        self.log_search_entries[tab] = log_search_entry

        return (tree, entry_output, txt_log, btn_frame, progress, status, 
                log_search_var, log_search_entry, log_prev_btn, log_next_btn, log_match_label)

    def build_user_tab(self, tab):
        (self.user_tree, self.user_output, self.user_log, btn_frame, self.user_progress, self.user_status,
         self.user_search_var, self.user_search_entry, self.user_prev_btn, self.user_next_btn, self.user_match_label,
         ) = self.build_common_layout(tab, "Select Excel file for User Generation:")
        self.init_tree_state(self.user_tree)
        tab.drop_target_register(DND_FILES)
        tab.dnd_bind("<<Drop>>", lambda e: self.import_files(self.root.tk.splitlist(e.data), self.user_tree, self.user_files, self.user_count,),)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=0, column=3, padx=(10, 5), pady=5, sticky="ns")
        self.user_add_btn = ttk.Button(button_frame, text="➕  Add Files", command=lambda: self.add_files(self.user_tree, self.user_files, self.user_count))
        self.user_add_btn.pack(fill="x", pady=(0, 10))
        self.user_remove_btn = ttk.Button(button_frame, text="➖ Remove Selected", command=lambda: self.remove_selected_file(self.user_tree, self.user_files, self.user_count))
        self.user_remove_btn.pack(fill="x", pady=(0, 10))
        self.user_clear_btn = ttk.Button(button_frame, text="❌  Clear All", command=lambda: self.clear_files(self.user_tree, self.user_files, self.user_count))
        self.user_clear_btn.pack(fill="x")
        self.user_count = ttk.Label(button_frame, text="Files Imported: 0", font=("Segoe UI", 9, "bold"))
        self.user_count.pack(pady=(15,0))
        self.user_generate_btn = ttk.Button(btn_frame, text="▶ Generate Script", command=lambda: self.run_in_thread(self.run_user_script))
        self.user_generate_btn.grid(row=0, column=0, padx=5)
        self.user_stop_btn = ttk.Button(btn_frame, text="⏹ Stop", command=self.stop_processing, state="disabled")
        self.user_stop_btn.grid(row=0, column=1, padx=5)
        self.user_copy_logs_btn = ttk.Button(btn_frame, text="📋 Copy Logs", command=lambda: self.copy_logs(self.user_log))
        self.user_copy_logs_btn.grid(row=0, column=2, padx=5)
        self.user_export_logs_btn = ttk.Button(btn_frame, text="💾 Export Logs", command=lambda: self.export_logs(self.user_log))
        self.user_export_logs_btn.grid(row=0, column=3, padx=5)
        self.user_clear_logs_btn = ttk.Button(btn_frame, text="🗑 Clear Logs", command=lambda: self.clear_logs(self.user_log))
        self.user_clear_logs_btn.grid(row=0, column=4, padx=5)
        self.user_export_template_btn = ttk.Button(btn_frame, text="📄 Export Template", command=lambda: export_template(USER_TEMPLATE_COLUMNS,USER_TEMPLATE_DATA, "UserTemplate.xlsx"))
        self.user_export_template_btn.grid(row=0, column=5, padx=5)
        self.user_controls = [self.user_add_btn,
                              self.user_remove_btn,
                              self.user_clear_btn,
                              self.user_generate_btn,
                              self.user_stop_btn,
                              self.user_copy_logs_btn,
                              self.user_export_logs_btn, 
                              self.user_clear_logs_btn,
                              self.user_export_template_btn]
        self.bind_treeview_shortcuts(self.user_tree, self.user_files, self.user_count)

    def build_address_tab(self, tab):
        (self.addr_tree, self.addr_output, self.addr_log, btn_frame, self.addr_progress, self.addr_status,
         self.addr_search_var, self.addr_search_entry, self.addr_prev_btn, self.addr_next_btn, self.addr_match_label,
         ) = self.build_common_layout(tab, "Select Excel file for Address Objects:")
        self.init_tree_state(self.addr_tree)
        tab.drop_target_register(DND_FILES)
        tab.dnd_bind("<<Drop>>", lambda e: self.import_files(self.root.tk.splitlist(e.data), self.addr_tree, self.addr_files, self.addr_count,),)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=0, column=3, padx=(10, 5), pady=5, sticky="ns")
        self.addr_add_btn = ttk.Button(button_frame, text="➕  Add Files", command=lambda: self.add_files(self.addr_tree, self.addr_files, self.addr_count))
        self.addr_add_btn.pack(fill="x", pady=(0, 10))
        self.addr_remove_btn = ttk.Button(button_frame, text="➖ Remove Selected", command=lambda: self.remove_selected_file(self.addr_tree, self.addr_files, self.addr_count))
        self.addr_remove_btn.pack(fill="x", pady=(0, 10))
        self.addr_clear_btn = ttk.Button(button_frame, text="❌  Clear All", command=lambda: self.clear_files(self.addr_tree, self.addr_files, self.addr_count))
        self.addr_clear_btn.pack(fill="x")
        self.addr_count = ttk.Label(button_frame, text="Files Imported: 0", font=("Segoe UI", 9, "bold"))
        self.addr_count.pack(pady=(15,0))
        self.addr_generate_btn = ttk.Button(btn_frame, text="▶ Generate Script", command=lambda: self.run_in_thread(self.run_address_script))
        self.addr_generate_btn.grid(row=0, column=0, padx=5)
        self.addr_stop_btn = ttk.Button(btn_frame, text="⏹ Stop", command=self.stop_processing, state="disabled")
        self.addr_stop_btn.grid(row=0, column=1, padx=5)
        self.addr_copy_logs_btn = ttk.Button(btn_frame, text="📋 Copy Logs", command=lambda: self.copy_logs(self.addr_log))
        self.addr_copy_logs_btn.grid(row=0, column=2, padx=5)
        self.addr_export_logs_btn = ttk.Button(btn_frame, text="💾 Export Logs", command=lambda: self.export_logs(self.addr_log))
        self.addr_export_logs_btn.grid(row=0, column=3, padx=5)
        self.addr_clear_logs_btn = ttk.Button(btn_frame, text="🗑 Clear Logs", command=lambda: self.clear_logs(self.addr_log))
        self.addr_clear_logs_btn.grid(row=0, column=4, padx=5)
        self.addr_export_template_btn = ttk.Button(btn_frame, text="📄 Export Template", command=lambda: export_template(ADDRESS_TEMPLATE_COLUMNS,ADDRESS_TEMPLATE_DATA, "AddressObjectTemplate.xlsx"))
        self.addr_export_template_btn.grid(row=0, column=5, padx=5)
        self.addr_controls = [self.addr_add_btn,
                              self.addr_remove_btn,
                              self.addr_clear_btn,
                              self.addr_generate_btn,
                              self.addr_stop_btn,
                              self.addr_copy_logs_btn,
                              self.addr_export_logs_btn, 
                              self.addr_clear_logs_btn,
                              self.addr_export_template_btn]
        self.bind_treeview_shortcuts(self.addr_tree, self.addr_files, self.addr_count)


    def build_service_tab(self, tab):
        (self.svc_tree, self.svc_output, self.svc_log, btn_frame, self.svc_progress, self.svc_status,
         self.svc_search_var, self.svc_search_entry, self.svc_prev_btn, self.svc_next_btn, self.svc_match_label,
         ) = self.build_common_layout(tab, "Select Excel file for Service Objects:")
        self.init_tree_state(self.svc_tree)
        tab.drop_target_register(DND_FILES)
        tab.dnd_bind("<<Drop>>", lambda e: self.import_files(self.root.tk.splitlist(e.data), self.svc_tree, self.svc_files, self.svc_count,),)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=0, column=3, padx=(10, 5), pady=5, sticky="ns")
        self.svc_add_btn = ttk.Button(button_frame, text="➕  Add Files", command=lambda: self.add_files(self.svc_tree, self.svc_files, self.svc_count))
        self.svc_add_btn.pack(fill="x", pady=(0, 10))
        self.svc_remove_btn = ttk.Button(button_frame, text="➖ Remove Selected", command=lambda: self.remove_selected_file(self.svc_tree, self.svc_files, self.svc_count))
        self.svc_remove_btn.pack(fill="x", pady=(0, 10))
        self.svc_clear_btn = ttk.Button(button_frame, text="❌  Clear All", command=lambda: self.clear_files(self.svc_tree, self.svc_files, self.svc_count))
        self.svc_clear_btn.pack(fill="x")
        self.svc_count = ttk.Label(button_frame, text="Files Imported: 0", font=("Segoe UI", 9, "bold"))
        self.svc_count.pack(pady=(15,0))
        self.svc_generate_btn = ttk.Button(btn_frame, text="▶ Generate Script", command=lambda: self.run_in_thread(self.run_service_script))
        self.svc_generate_btn.grid(row=0, column=0, padx=5)
        self.svc_stop_btn = ttk.Button(btn_frame, text="⏹ Stop", command=self.stop_processing, state="disabled")
        self.svc_stop_btn.grid(row=0, column=1, padx=5)
        self.svc_copy_logs_btn = ttk.Button(btn_frame, text="📋 Copy Logs", command=lambda: self.copy_logs(self.svc_log))
        self.svc_copy_logs_btn.grid(row=0, column=2, padx=5)
        self.svc_export_logs_btn = ttk.Button(btn_frame, text="💾 Export Logs", command=lambda: self.export_logs(self.svc_log))
        self.svc_export_logs_btn.grid(row=0, column=3, padx=5)
        self.svc_clear_logs_btn = ttk.Button(btn_frame, text="🗑 Clear Logs", command=lambda: self.clear_logs(self.svc_log))
        self.svc_clear_logs_btn.grid(row=0, column=4, padx=5)
        self.svc_export_template_btn = ttk.Button(btn_frame, text="📄 Export Template", command=lambda: export_template(SERVICE_TEMPLATE_COLUMNS,SERVICE_TEMPLATE_DATA, "ServiceObjectTemplate.xlsx"))
        self.svc_export_template_btn.grid(row=0, column=5, padx=5)
        self.svc_controls = [self.svc_add_btn,
                              self.svc_remove_btn,
                              self.svc_clear_btn,
                              self.svc_generate_btn,
                              self.svc_stop_btn,
                              self.svc_copy_logs_btn,
                              self.svc_export_logs_btn, 
                              self.svc_clear_logs_btn,
                              self.svc_export_template_btn]
        self.bind_treeview_shortcuts(self.svc_tree, self.svc_files, self.svc_count)


    def build_group_tab(self, tab):
        (self.grp_tree, self.grp_output, self.grp_log, btn_frame, self.grp_progress, self.grp_status,
         self.grp_search_var, self.grp_search_entry, self.grp_prev_btn, self.grp_next_btn, self.grp_match_label,
         ) = self.build_common_layout(tab, "Select Excel file for Group Generation:")
        self.init_tree_state(self.grp_tree)
        tab.drop_target_register(DND_FILES)
        tab.dnd_bind("<<Drop>>", lambda e: self.import_files(self.root.tk.splitlist(e.data), self.grp_tree, self.grp_files, self.grp_count,),)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=0, column=3, padx=(10, 5), pady=5, sticky="ns")
        self.grp_add_btn = ttk.Button(button_frame, text="➕  Add Files", command=lambda: self.add_files(self.grp_tree, self.grp_files, self.grp_count))
        self.grp_add_btn.pack(fill="x", pady=(0, 10))
        self.grp_remove_btn = ttk.Button(button_frame, text="➖ Remove Selected", command=lambda: self.remove_selected_file(self.grp_tree, self.grp_files, self.grp_count))
        self.grp_remove_btn.pack(fill="x", pady=(0, 10))
        self.grp_clear_btn = ttk.Button(button_frame, text="❌  Clear All", command=lambda: self.clear_files(self.grp_tree, self.grp_files, self.grp_count))
        self.grp_clear_btn.pack(fill="x")
        self.grp_count = ttk.Label(button_frame, text="Files Imported: 0", font=("Segoe UI", 9, "bold"))
        self.grp_count.pack(pady=(15,0))
        self.grp_generate_btn = ttk.Button(btn_frame, text="▶ Generate Script", command=lambda: self.run_in_thread(self.run_group_script))
        self.grp_generate_btn.grid(row=0, column=0, padx=5)
        self.grp_stop_btn = ttk.Button(btn_frame, text="⏹ Stop", command=self.stop_processing, state="disabled")
        self.grp_stop_btn.grid(row=0, column=1, padx=5)
        self.grp_copy_logs_btn = ttk.Button(btn_frame, text="📋 Copy Logs", command=lambda: self.copy_logs(self.grp_log))
        self.grp_copy_logs_btn.grid(row=0, column=2, padx=5)
        self.grp_export_logs_btn = ttk.Button(btn_frame, text="💾 Export Logs", command=lambda: self.export_logs(self.grp_log))
        self.grp_export_logs_btn.grid(row=0, column=3, padx=5)
        self.grp_clear_logs_btn = ttk.Button(btn_frame, text="🗑 Clear Logs", command=lambda: self.clear_logs(self.grp_log))
        self.grp_clear_logs_btn.grid(row=0, column=4, padx=5)
        self.grp_export_template_btn = ttk.Button(btn_frame, text="📄 Export Template", command=lambda: export_template(GROUP_TEMPLATE_COLUMNS,GROUP_TEMPLATE_DATA, "UserGroupTemplate.xlsx"))
        self.grp_export_template_btn.grid(row=0, column=5, padx=5)
        self.grp_controls = [self.grp_add_btn,
                              self.grp_remove_btn,
                              self.grp_clear_btn,
                              self.grp_generate_btn,
                              self.grp_stop_btn,
                              self.grp_copy_logs_btn,
                              self.grp_export_logs_btn, 
                              self.grp_clear_logs_btn,
                              self.grp_export_template_btn]
        self.bind_treeview_shortcuts(self.grp_tree, self.grp_files, self.grp_count)


    def build_dhcp_tab(self, tab):
        (self.dhcp_tree, self.dhcp_output, self.dhcp_log, btn_frame, self.dhcp_progress, self.dhcp_status,
         self.dhcp_search_var, self.dhcp_search_entry, self.dhcp_prev_btn, self.dhcp_next_btn, self.dhcp_match_label
         ) = self.build_common_layout(tab, "Select Excel/CSV file for DHCP Static Reservations:")
        self.init_tree_state(self.dhcp_tree)
        tab.drop_target_register(DND_FILES)
        tab.dnd_bind("<<Drop>>", lambda e: self.import_files(self.root.tk.splitlist(e.data), self.dhcp_tree, self.dhcp_files, self.dhcp_count,),)
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=0, column=3, padx=(10, 5), pady=5, sticky="ns")
        self.dhcp_add_btn = ttk.Button(button_frame, text="➕  Add Files", command=lambda: self.add_files(self.dhcp_tree, self.dhcp_files, self.dhcp_count))
        self.dhcp_add_btn.pack(fill="x", pady=(0, 10))
        self.dhcp_remove_btn = ttk.Button(button_frame, text="➖ Remove Selected", command=lambda: self.remove_selected_file(self.dhcp_tree, self.dhcp_files, self.dhcp_count))
        self.dhcp_remove_btn.pack(fill="x", pady=(0, 10))
        self.dhcp_clear_btn = ttk.Button(button_frame, text="❌  Clear All", command=lambda: self.clear_files(self.dhcp_tree, self.dhcp_files, self.dhcp_count))
        self.dhcp_clear_btn.pack(fill="x")
        self.dhcp_count = ttk.Label(button_frame, text="Files Imported: 0", font=("Segoe UI", 9, "bold"))
        self.dhcp_count.pack(pady=(15,0))
        self.dhcp_generate_btn = ttk.Button(btn_frame, text="▶ Generate Script", command=lambda: self.run_in_thread(self.run_dhcp_script))
        self.dhcp_generate_btn.grid(row=0, column=0, padx=5)
        self.dhcp_stop_btn = ttk.Button(btn_frame, text="⏹ Stop", command=self.stop_processing, state="disabled")
        self.dhcp_stop_btn.grid(row=0, column=1, padx=5)
        self.dhcp_copy_logs_btn = ttk.Button(btn_frame, text="📋 Copy Logs", command=lambda: self.copy_logs(self.dhcp_log))
        self.dhcp_copy_logs_btn.grid(row=0, column=2, padx=5)
        self.dhcp_export_logs_btn = ttk.Button(btn_frame, text="💾 Export Logs", command=lambda: self.export_logs(self.dhcp_log))
        self.dhcp_export_logs_btn.grid(row=0, column=3, padx=5)
        self.dhcp_clear_logs_btn = ttk.Button(btn_frame, text="🗑 Clear Logs", command=lambda: self.clear_logs(self.dhcp_log))
        self.dhcp_clear_logs_btn.grid(row=0, column=4, padx=5)
        self.dhcp_export_template_btn = ttk.Button(btn_frame, text="📄 Export Template", command=lambda: export_template(DHCP_TEMPLATE_COLUMNS, DHCP_TEMPLATE_DATA, "DHCP_Static_Template.xlsx"))
        self.dhcp_export_template_btn.grid(row=0, column=5, padx=5)
        self.dhcp_controls = [self.dhcp_add_btn,
                              self.dhcp_remove_btn,
                              self.dhcp_clear_btn,
                              self.dhcp_generate_btn,
                              self.dhcp_stop_btn,
                              self.dhcp_copy_logs_btn,
                              self.dhcp_export_logs_btn, 
                              self.dhcp_clear_logs_btn,
                              self.dhcp_export_template_btn]
        self.bind_treeview_shortcuts(self.dhcp_tree, self.dhcp_files, self.dhcp_count)


    def gui(self, func, *args, **kwargs):

        if threading.current_thread() is threading.main_thread():
            func(*args, **kwargs)
        else:
            self.gui_queue.put((func, args, kwargs))

    def process_gui_queue(self):
        while not self.gui_queue.empty():
            func, args, kwargs = self.gui_queue.get()
            func(*args, **kwargs)

        self.root.after(100, self.process_gui_queue)


    def thread_clear_files(self, tree, files, count):
        self.gui(self.clear_files, tree, files, count)

    def thread_startfile(self, path):
        self.gui(os.startfile, path)

    def thread_log(self, log_widget, message):
        self.gui(self.log, log_widget, message)

    def thread_info(self, title, message):
        self.gui(messagebox.showinfo, title, message)

    def thread_warning(self, title, message):
        self.gui(messagebox.showwarning, title, message)

    def thread_error(self, title, message):
        self.gui(messagebox.showerror, title, message)


    def on_tree_click_release(self, tree, event):
        # Ignore Shift-click because it should use the existing anchor
        if event.state & 0x0001:      # Shift
            return

        item = tree.focus()

        if item:
            self.set_tree_anchor(tree, item)

    def on_tree_ctrl_click(self, tree, event):
        item = tree.identify_row(event.y)

        if not item:
            return "break"

        selected = tree.selection()

        if item in selected:
            tree.selection_remove(item)
        else:
            tree.selection_add(item)

        tree.focus(item)
        self.set_tree_anchor(tree, item)
        tree.see(item)

        return "break"

    def init_tree_state(self, tree):
        self.tree_state[tree] = {
            "anchor": None
        }

    def get_tree_anchor(self, tree):
        return self.tree_state.setdefault(
            tree,
            {"anchor": None}
        )["anchor"]

    def set_tree_anchor(self, tree, item):
        self.tree_state.setdefault(
            tree,
            {"anchor": None}
        )["anchor"] = item

    def tree_move_focus(self, tree, target_index, extend=False):
        items = tree.get_children()

        if not items:
            return "break"

        # Clamp target index
        target_index = max(0, min(target_index, len(items) - 1))

        target = items[target_index]

        if extend:
            anchor = self.get_tree_anchor(tree)

            if anchor not in items or anchor is None:
                anchor = target
                self.set_tree_anchor(tree, anchor)

            anchor_index = items.index(anchor)

            start = min(anchor_index, target_index)
            end = max(anchor_index, target_index)

            tree.selection_set(items[start:end + 1])

        else:
            tree.selection_set(target)
            self.set_tree_anchor(tree, target)

        tree.focus(target)
        tree.see(target)

        return "break"

    def set_controls_state(self, controls, state):
        for widget in controls:
            try:
                widget.config(state=state)
            except:
                pass

    def tree_up(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        focus = tree.focus()

        if focus not in items:
            focus = items[0]
            self.set_tree_anchor(tree, focus)
            return self.tree_move_focus(tree, 0)

        index = items.index(focus)

        return self.tree_move_focus(tree, index - 1)

    def tree_down(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        focus = tree.focus()

        if focus not in items:
            focus = items[0]
            self.set_tree_anchor(tree, focus)
            return self.tree_move_focus(tree, 0)

        index = items.index(focus)

        return self.tree_move_focus(tree, index + 1)

    def tree_home(self, tree):
        return self.tree_move_focus(tree, 0)

    def tree_end(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        return self.tree_move_focus(tree, len(items) - 1)

    def tree_shift_down(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        focus = tree.focus()

        if focus not in items:
            return self.tree_move_focus(tree, 0)

        index = items.index(focus)

        # Extend selection downward
        return self.tree_move_focus(tree, index + 1, extend=True)

    def tree_shift_up(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        focus = tree.focus()

        if focus not in items:
            return self.tree_move_focus(tree, 0)

        index = items.index(focus)

        # Extend selection upward
        return self.tree_move_focus(tree, index - 1, extend=True)

    # Treeview Keyboard Shortcut Ctrl+A: Select All
    def tree_select_all(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        tree.selection_set(items)

        tree.focus(items[0])
        tree.see(items[0])

        self.set_tree_anchor(tree, items[0])

        return "break"

    def tree_shift_home(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        return self.tree_move_focus(tree, 0, extend=True)

    def tree_shift_end(self, tree):
        items = tree.get_children()

        if not items:
            return "break"

        return self.tree_move_focus(tree, len(items) - 1, extend=True)

    # Browse for output folder
    def browse(self, entry):
        path = filedialog.askdirectory()

        if path:
            entry.delete(0, END)
            entry.insert(0, path)

    # Update the file count label
    def update_file_count(self, label, file_list):
        label.config(text=f"Files Imported: {len(file_list)}")

    # Highlight the currently processing file in the Treeview
    def highlight_processing_file(self, tree, input_path):
        for item in tree.get_children():
            values = tree.item(item, "values")

            # values[1] is the Full Path column
            if values[1] == input_path:
                tree.selection_set(item)
                tree.focus(item)
                tree.see(item)
                break

    # Import files into the Treeview and update the file list
    def import_files(self, paths, tree, file_list, label):

        if not paths:
            return

        new_items = []

        for path in paths:

            path = os.path.normpath(path)

            if path in file_list:
                continue

            if not os.path.isfile(path):
                continue

            if not path.lower().endswith((".xlsx", ".xls", ".csv")):
                continue

            file_list.append(path)

            filename = os.path.basename(path)

            item = tree.insert(
                "",
                "end",
                values=(filename, path)
            )

            new_items.append(item)

        if new_items:
            tree.selection_set(new_items)
            tree.focus(new_items[-1])
            tree.see(new_items[-1])

        self.update_file_count(label, file_list)

    # Add files using a file dialog
    def add_files(self, tree, file_list, label):

        paths = filedialog.askopenfilenames(
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("CSV Files", "*.csv")
            ]
        )

        if paths:
            self.import_files(paths, tree, file_list, label)

    # Remove selected files from the Treeview and update the file list
    def remove_selected_file(self, tree, file_list, label):
        selected = tree.selection()
        if not selected:
            return

        for item in selected:
            values = tree.item(item, "values")
            full_path = values[1]  # Assuming the path is in the second column

            if full_path in file_list:
                file_list.remove(full_path)

            tree.delete(item)
            self.update_file_count(label, file_list)

    # Clear all files from the Treeview and update the file list
    def clear_files(self, tree, file_list, label):
        tree.delete(*tree.get_children())
        file_list.clear()
        self.update_file_count(label, file_list)

    def focus_log_search(self, event=None):
        current_tab = self.tabs.nametowidget(self.tabs.select())

        entry = self.log_search_entries.get(current_tab)

        if entry:
            entry.focus_set()
            entry.select_range(0, "end")
            entry.icursor("end")

        return "break"

    def show_current_match(self, txt_log, match_label):
        state = self.log_search_state[txt_log]

        if not state["matches"]:
            match_label.config(text="0 / 0")
            return

        # Remove previous current highlight
        txt_log.tag_remove("current_match", "1.0", "end")

        start, end = state["matches"][state["current"]]

        txt_log.tag_add("current_match", start, end)
        txt_log.see(start)

        match_label.config(
            text=f"{state['current'] + 1} / {len(state['matches'])}"
        )

    def highlight_log_matches(self, txt_log, search_var, match_label):
        # Remove previous highlights
        txt_log.tag_remove("search_match", "1.0", "end")
        txt_log.tag_remove("current_match", "1.0", "end")

        state = self.log_search_state[txt_log]
        state["matches"].clear()
        state["current"] = -1

        search_text = search_var.get()

        if not search_text:
            match_label.config(text="0 / 0")
            return

        start = "1.0"

        while True:
            pos = txt_log.search(
                search_text,
                start,
                stopindex="end",
                nocase=True
            )

            if not pos:
                break

            end = f"{pos}+{len(search_text)}c"

            txt_log.tag_add("search_match", pos, end)

            state["matches"].append((pos, end))

            start = end

        total = len(state["matches"])

        if total:
            state["current"] = 0
            self.show_current_match(txt_log, match_label)
        else:
            match_label.config(text="0 / 0")

    def goto_next_match(self, txt_log, match_label):
        state = self.log_search_state[txt_log]

        if not state["matches"]:
            return

        state["current"] += 1

        if state["current"] >= len(state["matches"]):
            state["current"] = 0

        self.show_current_match(txt_log, match_label)

    def goto_previous_match(self, txt_log, match_label):
        state = self.log_search_state[txt_log]

        if not state["matches"]:
            return

        state["current"] -= 1

        if state["current"] < 0:
            state["current"] = len(state["matches"]) - 1

        self.show_current_match(txt_log, match_label)

    # Stop the processing of files
    def stop_processing(self):
        self.stop_requested = True

    # Log messages to the Text widget
    def log(self, text_area, message):
        text_area.insert(END, message + "\n")
        text_area.see(END)

    # Clear the logs in the Text widget
    def clear_logs(self, text_area):
        text_area.delete(1.0, END)

    # Copy logs to the clipboard
    def copy_logs(self, text_area):
        logs = text_area.get("1.0", END).strip()
        if logs:
            text_area.clipboard_clear()
            text_area.clipboard_append(logs)
            text_area.update()
            self.thread_info("Copied", "Logs copied to clipboard.")

    # Export logs to a text file
    def export_logs(self, text_area):
        logs = text_area.get("1.0", END).strip()
        if not logs:
            self.thread_info("No Logs", "There are no logs to export.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    f.write(logs)
                self.thread_info("Success", f"Logs saved to:\n{file_path}")
            except Exception as e:
                self.thread_error("Error saving logs", str(e))

    # Thread-safe GUI updates
    def thread_status(self, label, text):
        self.gui(label.config, text=text)

    def thread_progress_value(self, progressbar, value):
        self.gui(progressbar.configure, value=value)

    def thread_progress_maximum(self, progressbar, maximum):
        self.gui(progressbar.configure, maximum=maximum)

    def thread_highlight(self, tree, path):
        self.gui(self.highlight_processing_file, tree, path)

    def run_in_thread(self, target):
        threading.Thread(target=target, daemon=True).start()

    # Run the user script processing in a separate thread
    def run_user_script(self):
        if not self.user_files:
            self.thread_error("Error", "Please add at least one Excel file.")
            return

        output_folder = self.user_output.get().strip()

        if not output_folder or not os.path.isdir(output_folder):
            self.thread_error("Error", "Please select a valid output folder.")
            return

        self.stop_requested = False

        total_files = len(self.user_files)

        self.thread_progress_maximum(self.user_progress, total_files)
        self.thread_progress_value(self.user_progress, 0)

        self.thread_status(self.user_status, "Starting...")

        processed_files = 0
        failed_files = 0
        processed_paths = []

        self.set_controls_state(self.user_controls, "disabled")
        self.user_stop_btn.config(state="normal")

        for input_path in self.user_files:
            if self.stop_requested:
                break
            self.thread_highlight(self.user_tree, input_path)
            try:
                df = pd.read_excel(input_path)
                df.columns = [c.strip() for c in df.columns]
                in_fname = os.path.basename(input_path)
                if 'Username' not in df.columns:
                    self.thread_log(self.user_log, f"[ERROR] {in_fname}: Missing 'Username' column")
                    failed_files += 1
                    continue
                ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                ts_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fname = os.path.splitext(in_fname)[0]
                out_fname = f"{fname}_LocalUsers_Script_{ts_file}.txt"
                full_out = os.path.join(output_folder, out_fname)
                self.thread_log(self.user_log, f"\nLogs for {in_fname} [{ts_log}]")
                self.thread_log(self.user_log, "-" * 70)
                commands, logs = generate_full_script(df)

                # Write logs
                for msg in logs:
                    self.thread_log(self.user_log, msg)

                # Write output file
                with open(full_out, 'w', encoding="utf-8") as f:
                    if commands:
                        f.write("user local\n")

                        for cmd in commands:
                            f.write(cmd + "\n")

                processed_files += 1
                processed_paths.append(input_path)

                self.thread_status(self.user_status, f"Processing: {in_fname} {processed_files}/{total_files}")
                self.thread_progress_value(self.user_progress, processed_files)

            except Exception as e:
                self.thread_log(self.user_log, f"[ERROR] {in_fname}: {e}")

        if self.stop_requested:

            # Remove successfully processed files
            for path in processed_paths:

                if path in self.user_files:
                    self.user_files.remove(path)

                for item in self.user_tree.get_children():
                    values = self.user_tree.item(item, "values")

                    if values[1] == path:
                        self.user_tree.delete(item)
                        break

            self.update_file_count(self.user_count, self.user_files)

            self.thread_status(self.user_status, "Stopped")

            self.thread_info(
                "Processing Stopped",
                f"✔ Successful: {processed_files}\n"
                f"✖ Failed: {failed_files}\n"
                f"⏭ Remaining: {len(self.user_files)}\n\n"
                f"Processed files have been removed from the list."
            )

            self.set_controls_state(self.user_controls, "normal")
            self.user_stop_btn.config(state="disabled")

            self.thread_startfile(output_folder)

            return

        if processed_files > 0:
            self.thread_status(self.user_status, "Completed")
            self.thread_info("Completed", f"✔ Successful: {processed_files}\n"f"✖ Failed: {failed_files}\n\n"f"Output Folder:\n{output_folder}")
            self.set_controls_state(self.user_controls, "normal")
            self.user_stop_btn.config(state="disabled")
            if failed_files == 0:
                self.thread_clear_files(self.user_tree, self.user_files, self.user_count)
            self.thread_startfile(output_folder)
        else:
            self.thread_status(self.user_status, "Failed")
            self.thread_warning("Completed", "No files were processed successfully.\n\n""Please check the log window for details.")
            self.set_controls_state(self.user_controls, "normal")
            self.user_stop_btn.config(state="disabled")


    def run_address_script(self):
        if not self.addr_files:
            self.thread_error("Error", "Please add at least one Excel file.")
            return

        output_folder = self.addr_output.get().strip()

        if not output_folder or not os.path.isdir(output_folder):
            self.thread_error("Error", "Please select a valid output folder.")
            return

        self.stop_requested = False

        total_files = len(self.addr_files)
        self.thread_progress_maximum(self.addr_progress, total_files)
        self.thread_progress_value(self.addr_progress, 0)
        self.thread_status(self.addr_status, "Starting...")

        processed_files = 0
        failed_files = 0
        processed_paths = []

        self.set_controls_state(self.addr_controls, "disabled")
        self.addr_stop_btn.config(state="normal")

        for input_path in self.addr_files:
            if self.stop_requested:
                break
            self.thread_highlight(self.addr_tree, input_path)
            try:
                df = pd.read_excel(input_path)
                df.columns = [c.strip() for c in df.columns]
                in_fname = os.path.basename(input_path)
                if not {'Name', 'Address', 'Zone'}.issubset(df.columns):
                    self.thread_log(self.addr_log, f"[ERROR] {in_fname}: Excel file must contain columns: Name, Address, Zone")
                    failed_files += 1
                    continue
                ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                ts_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fname = os.path.splitext(in_fname)[0]
                out_fname = f"{fname}_AddressObjects_{ts_file}.txt"
                full_out = os.path.join(output_folder, out_fname)
                self.thread_log(self.addr_log, f"\nLogs for {in_fname} [{ts_log}]")
                self.thread_log(self.addr_log, "-"*70)
                address_groups = {}  # {group_name: set((object_name, object_type))}
                with open(full_out, 'w', encoding="utf-8") as f:
                    for i, row in df.iterrows():
                        name = str(row['Name']).strip() if pd.notna(row['Name']) else ''
                        addr = str(row['Address']).strip() if pd.notna(row['Address']) else ''
                        zone = str(row['Zone']).strip() if pd.notna(row['Zone']) else ''
                        grp_raw = str(row['Group']).strip() if 'Group' in row and pd.notna(row['Group']) else ''
                        groups = [g.strip() for g in grp_raw.split(',') if g.strip()]
                        cmds, logs = generate_address_command(name, addr, zone)
                        for msg in logs: self.thread_log(self.addr_log, msg)
                        for cmd in cmds:
                            f.write(cmd + "\n")
                        if cmds: 
                            f.write("\n")
                            addr_type = detect_type(addr)
                            obj_type = 'ipv4' if addr_type in ('host', 'network', 'range') else addr_type
                            obj_name = name or addr

                            for grp in groups:
                                address_groups.setdefault(grp, set()).add((obj_name, obj_type))

                    for grp, members in address_groups.items():
                        first_obj_type = next(iter(members))[1]
                        grp_type = 'ipv4' if all(obj_type == 'ipv4' for _, obj_type in members) else 'ipv6'
                        quoted_grp = f'"{grp}"' if needs_quotes(grp) else grp

                        f.write(f'address-group {grp_type} {quoted_grp}\n')

                        for obj_name, obj_type in sorted(members):
                            quoted_obj = f'"{obj_name}"' if needs_quotes(obj_name) else obj_name
                            f.write(f'address-object {obj_type} {quoted_obj}\n')
                        f.write("exit\n\n")

                processed_files += 1
                processed_paths.append(input_path)

                self.thread_status(self.addr_status, f"Processing: {in_fname} {processed_files}/{total_files}")
                self.thread_progress_value(self.addr_progress, processed_files)

            except Exception as e:
                self.thread_log(self.addr_log, f"[ERROR] {in_fname}: {e}")

        if self.stop_requested:

            # Remove successfully processed files
            for path in processed_paths:

                if path in self.addr_files:
                    self.addr_files.remove(path)

                for item in self.addr_tree.get_children():
                    values = self.addr_tree.item(item, "values")

                    if values[1] == path:
                        self.addr_tree.delete(item)
                        break

            self.update_file_count(self.addr_count, self.addr_files)

            self.thread_status(self.addr_status, "Stopped")

            self.thread_info(
                "Processing Stopped",
                f"✔ Successful: {processed_files}\n"
                f"✖ Failed: {failed_files}\n"
                f"⏭ Remaining: {len(self.addr_files)}\n\n"
                f"Processed files have been removed from the list."
            )

            self.set_controls_state(self.addr_controls, "normal")
            self.addr_stop_btn.config(state="disabled")

            self.thread_startfile(output_folder)

            return

        if processed_files > 0:
            self.thread_status(self.addr_status, "Completed")
            self.thread_info("Completed", f"✔ Successful: {processed_files}\n"f"✖ Failed: {failed_files}\n\n"f"Output Folder:\n{output_folder}")
            self.set_controls_state(self.addr_controls, "normal")
            self.addr_stop_btn.config(state="disabled")
            if failed_files == 0:
                self.thread_clear_files(self.addr_tree, self.addr_files, self.addr_count)
            self.thread_startfile(output_folder)
        else:
            self.thread_status(self.addr_status, "Failed")
            self.thread_warning("Completed", "No files were processed successfully.\n\n""Please check the log window for details.")
            self.set_controls_state(self.addr_controls, "normal")
            self.addr_stop_btn.config(state="disabled")


    def run_service_script(self):
        if not self.svc_files:
            self.thread_error("Error", "Please add at least one Excel file.")
            return
        
        output_folder = self.svc_output.get().strip()

        if not output_folder or not os.path.isdir(output_folder):
            self.thread_error("Error", "Please select a valid output folder.")
            return

        self.stop_requested = False

        total_files = len(self.svc_files)
        self.thread_progress_maximum(self.svc_progress, total_files)
        self.thread_progress_value(self.svc_progress, 0)
        self.thread_status(self.svc_status, "Starting...")

        processed_files = 0
        failed_files = 0
        processed_paths = []

        self.set_controls_state(self.svc_controls, "disabled")
        self.svc_stop_btn.config(state="normal")

        for input_path in self.svc_files:
            if self.stop_requested:
                break
            self.thread_highlight(self.svc_tree, input_path)
            try:
                df = pd.read_excel(input_path)
                df.columns = [c.strip() for c in df.columns]
                required_cols = {'Name', 'Protocol', 'Port'}
                in_fname = os.path.basename(input_path)
                if not required_cols.issubset(df.columns):
                    self.thread_log(self.svc_log, f"[ERROR] {in_fname}: Excel file must contain columns: {', '.join(required_cols)}")
                    failed_files += 1
                    continue
                ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                ts_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fname = os.path.splitext(in_fname)[0]
                out_fname = f"{fname}_ServiceObjects_{ts_file}.txt"
                full_out = os.path.join(output_folder, out_fname)
                self.thread_log(self.svc_log, f"\nLogs for {in_fname} [{ts_log}]")
                self.thread_log(self.svc_log, "-"*70)
                service_groups = {}
                with open(full_out, 'w', encoding="utf-8") as f:
                    for i, row in df.iterrows():
                        name = str(row['Name']).strip() if pd.notna(row['Name']) else ''
                        protocol = str(row['Protocol']).strip() if pd.notna(row['Protocol']) else ''
                        port = str(row['Port']).strip() if pd.notna(row['Port']) else ''
                        grp_raw = str(row['Group']).strip() if 'Group' in row and pd.notna(row['Group']) else ''
                        groups = [g.strip() for g in grp_raw.split(',') if g.strip()]
                        cmds, logs = generate_service_command(name, protocol, port)
                        for msg in logs: self.thread_log(self.svc_log, msg)
                        for cmd in cmds:
                            f.write(cmd + "\n")
                        if cmds:
                            f.write("\n")
                            for grp in groups:
                                service_groups.setdefault(grp, set()).add(name)
                    for grp, members in service_groups.items():
                        quoted_grp = f'"{grp}"' if needs_quotes(grp) else grp
                        f.write(f'service-group {quoted_grp}\n')

                        for svc in sorted(members):
                            quoted_svc = f'"{svc}"' if needs_quotes(svc) else svc
                            f.write(f'service-object {quoted_svc}\n')
                        f.write("exit\n\n")

                processed_files += 1
                processed_paths.append(input_path)

                self.thread_status(self.svc_status, f"Processing: {in_fname} {processed_files}/{total_files}")
                self.thread_progress_value(self.svc_progress, processed_files)

            except Exception as e:
                self.thread_log(self.svc_log, f"[ERROR] {in_fname}: {e}")

        if self.stop_requested:

            # Remove successfully processed files
            for path in processed_paths:

                if path in self.svc_files:
                    self.svc_files.remove(path)

                for item in self.svc_tree.get_children():
                    values = self.svc_tree.item(item, "values")

                    if values[1] == path:
                        self.svc_tree.delete(item)
                        break

            self.update_file_count(self.svc_count, self.svc_files)

            self.thread_status(self.svc_status, "Stopped")

            self.thread_info(
                "Processing Stopped",
                f"✔ Successful: {processed_files}\n"
                f"✖ Failed: {failed_files}\n"
                f"⏭ Remaining: {len(self.svc_files)}\n\n"
                f"Processed files have been removed from the list."
            )

            self.set_controls_state(self.svc_controls, "normal")
            self.svc_stop_btn.config(state="disabled")

            self.thread_startfile(output_folder)

            return

        if processed_files > 0:
            self.thread_status(self.svc_status, "Completed")
            self.thread_info("Completed", f"✔ Successful: {processed_files}\n"f"✖ Failed: {failed_files}\n\n"f"Output Folder:\n{output_folder}")
            self.set_controls_state(self.svc_controls, "normal")
            self.svc_stop_btn.config(state="disabled")
            if failed_files == 0:
                self.thread_clear_files(self.svc_tree, self.svc_files, self.svc_count)
            self.thread_startfile(output_folder)
        else:
            self.thread_status(self.svc_status, "Failed")
            self.thread_warning("Completed", "No files were processed successfully.\n\n""Please check the log window for details.")
            self.set_controls_state(self.svc_controls, "normal")
            self.svc_stop_btn.config(state="disabled")


    def run_group_script(self):
        if not self.grp_files:
            self.thread_error("Error", "Please add at least one Excel file.")
            return

        output_folder = self.grp_output.get().strip()

        if not output_folder or not os.path.isdir(output_folder):
            self.thread_error("Error", "Please select a valid output folder.")
            return

        self.stop_requested = False

        total_files = len(self.grp_files)
        self.thread_progress_maximum(self.grp_progress, total_files)
        self.thread_progress_value(self.grp_progress, 0)
        self.thread_status(self.grp_status, "Starting...")

        processed_files = 0
        failed_files = 0
        processed_paths = []

        self.set_controls_state(self.grp_controls, "disabled")
        self.grp_stop_btn.config(state="normal")

        for input_path in self.grp_files:
            if self.stop_requested:
                break
            self.thread_highlight(self.grp_tree, input_path)
            try:
                df = pd.read_excel(input_path)
                df.columns = [c.strip() for c in df.columns]
                in_fname = os.path.basename(input_path)
                if 'Group' not in df.columns:
                    self.thread_log(self.grp_log, f"[ERROR] {in_fname}: Excel file must contain 'Group' column")
                    failed_files += 1
                    continue
                ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                ts_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fname = os.path.splitext(in_fname)[0]
                out_fname = f"{fname}_LocalGroups_{ts_file}.txt"
                full_out = os.path.join(output_folder, out_fname)
                self.thread_log(self.grp_log, f"\nLogs for {in_fname} [{ts_log}]")
                self.thread_log(self.grp_log, "-" * 70)
                with open(full_out, 'w', encoding="utf-8") as f:
                    for i, row in df.iterrows():
                        group_name = str(row['Group']).strip() if pd.notna(row['Group']) else ''
                        members = str(row['Member']).strip() if 'Member' in row and pd.notna(row['Member']) else ''
                        cmds, logs = generate_group_command(group_name, members, row_number=i+2)
                        for msg in logs: self.thread_log(self.grp_log, msg)
                        for cmd in cmds:
                            f.write(cmd + "\n")
                        if cmds: f.write("\n")

                processed_files += 1
                processed_paths.append(input_path)

                self.thread_status(self.grp_status, f"Processing: {in_fname} {processed_files}/{total_files}")
                self.thread_progress_value(self.grp_progress, processed_files)

            except Exception as e:
                self.thread_log(self.grp_log, f"[ERROR] {in_fname}: {e}")

        if self.stop_requested:

            # Remove successfully processed files
            for path in processed_paths:

                if path in self.grp_files:
                    self.grp_files.remove(path)

                for item in self.grp_tree.get_children():
                    values = self.grp_tree.item(item, "values")

                    if values[1] == path:
                        self.grp_tree.delete(item)
                        break

            self.update_file_count(self.grp_count, self.grp_files)

            self.thread_status(self.grp_status, "Stopped")

            self.thread_info(
                "Processing Stopped",
                f"✔ Successful: {processed_files}\n"
                f"✖ Failed: {failed_files}\n"
                f"⏭ Remaining: {len(self.grp_files)}\n\n"
                f"Processed files have been removed from the list."
            )

            self.set_controls_state(self.grp_controls, "normal")
            self.grp_stop_btn.config(state="disabled")

            self.thread_startfile(output_folder)

            return

        if processed_files > 0:
            self.thread_status(self.grp_status, "Completed")
            self.thread_info("Completed", f"✔ Successful: {processed_files}\n"f"✖ Failed: {failed_files}\n\n"f"Output Folder:\n{output_folder}")
            self.set_controls_state(self.grp_controls, "normal")
            self.grp_stop_btn.config(state="disabled")
            if failed_files == 0:
                self.thread_clear_files(self.grp_tree, self.grp_files, self.grp_count)
            self.thread_startfile(output_folder)
        else:
            self.thread_status(self.grp_status, "Failed")
            self.thread_warning("Completed", "No files were processed successfully.\n\n""Please check the log window for details.")
            self.set_controls_state(self.grp_controls, "normal")
            self.grp_stop_btn.config(state="disabled")


    def run_dhcp_script(self):
        if not self.dhcp_files:
            self.thread_error("Error", "Please add at least one Excel or CSV file.")
            return

        output_folder = self.dhcp_output.get().strip()

        if not os.path.isdir(output_folder):
            self.thread_error("Error", "Please select a valid output folder.")
            return

        self.stop_requested = False

        total_files = len(self.dhcp_files)
        self.thread_progress_maximum(self.dhcp_progress, total_files)
        self.thread_progress_value(self.dhcp_progress, 0)
        self.thread_status(self.dhcp_status, "Starting...")

        processed_files = 0
        failed_files = 0
        processed_paths = []

        self.set_controls_state(self.dhcp_controls, "disabled")
        self.dhcp_stop_btn.config(state="normal")

        for input_path in self.dhcp_files:
            if self.stop_requested:
                break
            self.thread_highlight(self.dhcp_tree, input_path)
            try:
                if input_path.lower().endswith((".xls", ".xlsx")):
                    df = pd.read_excel(input_path)
                else:
                    df = pd.read_csv(input_path)

                df.columns = [c.strip() for c in df.columns]

                ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
                ts_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                fname = os.path.splitext(os.path.basename(input_path))[0]
                in_fname = os.path.basename(input_path)

                out_file = os.path.join(
                    output_folder,
                    f"{fname}_DHCP_Static_{ts_file}.txt"
                )

                self.thread_log(self.dhcp_log, f"\nLogs for {in_fname} [{ts_log}]")
                self.thread_log(self.dhcp_log, "-" * 70)

                log_lines = []
                script = generate_dhcp_cli_script(df, log_lines)

                with open(out_file, "w") as f:
                    f.write(script)

                for line in log_lines:
                    self.thread_log(self.dhcp_log, line)

                processed_files += 1
                processed_paths.append(input_path)

                self.thread_status(self.dhcp_status, f"Processing: {in_fname} {processed_files}/{total_files}")
                self.thread_progress_value(self.dhcp_progress, processed_files)

            except Exception as e:
                self.thread_error("Error", f"[ERROR] {in_fname}: {e}")
                failed_files += 1
                continue

        if self.stop_requested:

            # Remove successfully processed files
            for path in processed_paths:

                if path in self.dhcp_files:
                    self.dhcp_files.remove(path)

                for item in self.dhcp_tree.get_children():
                    values = self.dhcp_tree.item(item, "values")

                    if values[1] == path:
                        self.dhcp_tree.delete(item)
                        break

            self.update_file_count(self.dhcp_count, self.dhcp_files)

            self.thread_status(self.dhcp_status, "Stopped")

            self.thread_info(
                "Processing Stopped",
                f"✔ Successful: {processed_files}\n"
                f"✖ Failed: {failed_files}\n"
                f"⏭ Remaining: {len(self.dhcp_files)}\n\n"
                f"Processed files have been removed from the list."
            )

            self.set_controls_state(self.dhcp_controls, "normal")
            self.dhcp_stop_btn.config(state="disabled")

            self.thread_startfile(output_folder)

            return

        if processed_files > 0:
            self.thread_status(self.dhcp_status, "Completed")
            self.thread_info("Completed", f"✔ Successful: {processed_files}\n"f"✖ Failed: {failed_files}\n\n"f"Output Folder:\n{output_folder}")
            self.set_controls_state(self.dhcp_controls, "normal")
            self.dhcp_stop_btn.config(state="disabled")
            if failed_files == 0:
                self.thread_clear_files(self.dhcp_tree, self.dhcp_files, self.dhcp_count)
            self.thread_startfile(output_folder)
        else:
            self.thread_status(self.dhcp_status, "Failed")
            self.thread_warning("Completed", "No files were processed successfully.\n\n" "Please check the log window for details.")
            self.set_controls_state(self.dhcp_controls, "normal")
            self.dhcp_stop_btn.config(state="disabled")

# Confirmation before exit
def confirm_exit():
    if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
        root.quit()

# About Software
def show_about():
    messagebox.showinfo("About", f"SonicWallTool\n\nVersion: {VERSION}\n\nDeveloped for internal use only.")

# Calling
if __name__ == "__main__":
    root = TkinterDnD.Tk()

    # Create menu bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # File menu
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Exit", command=confirm_exit)

    # Make it accessible in other functions
    EXIT_MENU_INDEX = 0

    # Help menu
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Check for Updates", command=lambda: check_for_updates(auto=False))
    help_menu.add_separator()
    help_menu.add_command(label="About", command=show_about)

    # Launch main app
    app = SonicToolApp(root)

    # Automatically check updates at startup
    # root.after(1000, lambda: check_for_updates(auto=True))

    root.mainloop()
